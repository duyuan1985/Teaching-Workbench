"""
教学档案工作台 - FastAPI 后端

将原 app.py 的服务端渲染改为 REST API，
前端由 Vue 3 SPA 接管。
"""

import json
import hashlib
import os
import subprocess
import sys
import threading
from datetime import date
from pathlib import Path
from typing import Any, List, Optional

from fastapi import FastAPI, Form, UploadFile, File, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

import store
from assessment_scheme import ensure_scheme, get_scheme
from catalog import catalog_from_arrangement
from curriculum_review import build_curriculum_review
from document_generator import generate_offering_documents
from foundation_pipeline import rebuild_generation_foundation
from grade_analysis import generate_grade_analysis
from importers import import_calendar_week_dates, import_progress_table, import_school_calendar, import_teaching_arrangement, rebuild_schedule
from ai.ai_router import current_model, current_url, installed_models, ollama_available
from listening_records import generate_listening_record, templates as listening_templates
from resource_indexer import build_resource_index
from resource_analyzer import analyze_offering_resources
from task_builder import build_tasks
from template_analyzer import analyze_template
from contract_parser import parse_contract, scan_templates
from training_materials import generate_training_materials
from validators import generation_readiness, validate
from config import DEFAULTS, ALLOWED_UPLOAD_EXTENSIONS, MAX_UPLOAD_SIZE, HTTP_HOST, HTTP_PORT, CORS_ORIGINS, TEMPLATE_DIR
from content_updater import analyze_content_updates, format_updates_for_prompt

app = FastAPI(title="教学档案工作台 API", version="2.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_methods=["GET", "POST", "PUT", "DELETE"],
    allow_headers=["*"],
)

# 确保数据库已初始化
store.initialize()


class AppError(Exception):
    """业务错误，返回400"""
    pass


class NotFoundError(Exception):
    """资源不存在，返回404"""
    pass


class ForbiddenError(Exception):
    """权限不足，返回403"""
    pass


@app.exception_handler(AppError)
async def app_error_handler(request, exc):
    return JSONResponse(status_code=400, content={"ok": False, "error": {"code": "BUSINESS_ERROR", "message": str(exc)}})


@app.exception_handler(NotFoundError)
async def not_found_error_handler(request, exc):
    return JSONResponse(status_code=404, content={"ok": False, "error": {"code": "NOT_FOUND", "message": str(exc)}})


@app.exception_handler(ForbiddenError)
async def forbidden_error_handler(request, exc):
    return JSONResponse(status_code=403, content={"ok": False, "error": {"code": "FORBIDDEN", "message": str(exc)}})


@app.exception_handler(Exception)
async def generic_error_handler(request, exc):
    return JSONResponse(status_code=500, content={"ok": False, "error": {"code": "INTERNAL_ERROR", "message": "服务器内部错误"}})


# ============================================================
# 通用辅助
# ============================================================

def esc(text: Any) -> str:
    """HTML 转义（部分接口可能需要返回已转义文本）"""
    import html as _html
    return _html.escape(str(text)) if text is not None else ""


def parse_json_field(value: str, default=None):
    if not value:
        return default
    try:
        return json.loads(value)
    except (json.JSONDecodeError, TypeError):
        return default


ALLOWED_UPLOAD_EXTS = set(ALLOWED_UPLOAD_EXTENSIONS)


async def save_upload_file(upload: UploadFile, prefix: str, offering_id: int) -> Path:
    """安全保存上传文件到临时目录"""
    import tempfile, uuid
    ext = Path(upload.filename or "").suffix.lower()
    if ext not in ALLOWED_UPLOAD_EXTS:
        raise HTTPException(400, f"不支持的文件类型: {ext}，仅支持 {', '.join(ALLOWED_UPLOAD_EXTENSIONS)}")
    content = await upload.read()
    if len(content) > MAX_UPLOAD_SIZE:
        raise HTTPException(400, f"文件大小超过限制({MAX_UPLOAD_SIZE // 1024 // 1024}MB)")
    safe_name = f"{prefix}_{offering_id}_{uuid.uuid4().hex[:8]}{ext}"
    tmp_path = Path(tempfile.gettempdir()) / safe_name
    with open(tmp_path, "wb") as f:
        f.write(content)
    return tmp_path


# ============================================================
# 课程实例
# ============================================================

@app.get("/api/offerings")
def list_offerings(q: str = "", term: str = ""):
    sql = "SELECT * FROM offerings WHERE 1=1"
    params = []
    if q:
        sql += " AND (course_name LIKE ? OR course_code LIKE ? OR major LIKE ? OR teaching_class LIKE ?)"
        kw = f"%{q}%"
        params += [kw, kw, kw, kw]
    if term:
        sql += " AND term=?"
        params.append(term)
    sql += " ORDER BY term DESC, course_name, major, id DESC"
    return store.rows(sql, params)


@app.get("/api/offerings/stats")
def offerings_stats():
    offerings = store.rows("SELECT * FROM offerings ORDER BY term DESC, course_name, major, id DESC")
    terms = set(o["term"] for o in offerings if o.get("term"))
    total_hours = sum(o.get("total_hours", 0) or 0 for o in offerings)
    return {
        "total": len(offerings),
        "terms": len(terms),
        "total_hours": total_hours,
    }


@app.get("/api/offerings/{offering_id}")
def get_offering(offering_id: int):
    row = store.rows(
        """SELECT o.*,
            (SELECT COUNT(*) FROM tasks WHERE offering_id=o.id) AS task_count,
            (SELECT COUNT(*) FROM sessions WHERE offering_id=o.id AND status='已确认') AS session_count,
            (SELECT COUNT(*) FROM curriculum_units WHERE offering_id=o.id AND approval_status='已确认') AS unit_count,
            (SELECT COUNT(*) FROM resource_items WHERE offering_id=o.id) AS resource_count,
            (SELECT COUNT(*) FROM course_content_models WHERE offering_id=o.id) AS model_count,
            (SELECT COUNT(*) FROM authored_sections WHERE offering_id=o.id) AS draft_count,
            (SELECT COUNT(*) FROM generated_documents WHERE offering_id=o.id) AS doc_count,
            (SELECT COUNT(*) FROM template_files WHERE offering_id=o.id) AS template_file_count,
            (SELECT COUNT(*) FROM source_files WHERE offering_id=o.id) AS source_file_count
           FROM offerings o WHERE o.id=?""",
        (offering_id,),
    )
    if not row:
        raise HTTPException(404, "课程实例不存在")
    r = row[0]
    offering = {k: r[k] for k in r.keys() if not k.endswith("_count")}

    workflow = [
        {"step": 1, "name": "基本信息", "done": True},
        {"step": 2, "name": "排课确认", "done": r["session_count"] > 0},
        {"step": 3, "name": "资源索引", "done": r["resource_count"] > 0},
        {"step": 4, "name": "蓝本审查", "done": r["unit_count"] > 0},
        {"step": 5, "name": "内容生成", "done": r["draft_count"] > 0},
        {"step": 6, "name": "文档输出", "done": r["doc_count"] > 0},
    ]
    current_step = 1
    for w in workflow:
        if w["done"]:
            current_step = w["step"] + 1

    return {
        "offering": offering,
        "workflow": workflow,
        "current_step": min(current_step, 6),
        "counts": {
            "tasks": r["task_count"],
            "sessions": r["session_count"],
            "units": r["unit_count"],
            "resources": r["resource_count"],
            "models": r["model_count"],
            "drafts": r["draft_count"],
            "documents": r["doc_count"],
            "template_files": r["template_file_count"],
            "source_files": r["source_file_count"],
        },
    }


@app.get("/api/offerings/{offering_id}/dirty-flags")
def get_dirty_flags(offering_id: int):
    _get_offering(offering_id)
    flags = store.get_dirty_flags_with_meta(offering_id)
    active_flags = [f for f in flags if f["active"]]
    recommended = None
    if active_flags:
        priority = ["resources", "review", "tasks", "templates", "foundation", "basic_info", "schedule", "teacher_name"]
        for flag in priority:
            match = next((f for f in active_flags if f["flag"] == flag), None)
            if match:
                recommended = match
                break
    return {"flags": flags, "active_count": len(active_flags), "recommended": recommended}


class OfferingCreate(BaseModel):
    term: str
    course_name: str
    course_code: str = ""
    major: str
    teaching_class: str = ""
    course_nature: str = ""
    course_type: str = ""
    assessment_type: str = ""
    assessment_method: str = ""
    credits: float = 0
    total_hours: int = 0
    weekly_hours: int = 0
    textbook_version: str = ""
    textbook_path: str = ""
    template_path: str = ""
    schedule_path: str = ""
    notes: str = ""
    offering_kind: str = "普通课程"


@app.post("/api/offerings")
def create_offering(body: OfferingCreate):
    data = body.model_dump()
    with store.connect() as db:
        cursor = db.execute(
            """INSERT INTO offerings
            (term,course_name,course_code,major,teaching_class,course_nature,course_type,
             assessment_type,assessment_method,credits,total_hours,weekly_hours,
             textbook_version,textbook_path,template_path,schedule_path,notes,offering_kind)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (data["term"], data["course_name"], data["course_code"], data["major"],
             data["teaching_class"], data["course_nature"], data["course_type"],
             data["assessment_type"], data["assessment_method"], data["credits"],
             data["total_hours"], data["weekly_hours"], data["textbook_version"],
             data["textbook_path"], data["template_path"], data["schedule_path"],
             data["notes"], data["offering_kind"]),
        )
        db.commit()
        return {"id": cursor.lastrowid}


@app.get("/api/terms")
def list_terms():
    rows = store.rows("SELECT DISTINCT term FROM offerings WHERE term<>'' ORDER BY term DESC")
    return [r["term"] for r in rows]


class OfferingUpdate(BaseModel):
    course_name: str
    course_code: str = ""
    major: str = ""
    teaching_class: str = ""
    course_nature: str = ""
    course_type: str = ""
    assessment_type: str = ""
    assessment_method: str = ""
    credits: float = 0
    total_hours: int = 0
    weekly_hours: int = 0
    textbook_version: str = ""
    notes: str = ""


@app.put("/api/offerings/{offering_id}")
def update_offering(offering_id: int, body: OfferingUpdate):
    _get_offering(offering_id)
    if not body.course_name.strip():
        raise AppError("课程名不能为空")
    data = body.model_dump()
    fields = [
        "course_name", "course_code", "major", "teaching_class", "course_nature",
        "course_type", "assessment_type", "assessment_method", "credits",
        "total_hours", "weekly_hours", "textbook_version", "notes",
    ]
    with store.connect() as db:
        db.execute(
            f"UPDATE offerings SET {','.join(f + '=?' for f in fields)} WHERE id=?",
            [data[f] for f in fields] + [offering_id],
        )
        db.commit()
    store.mark_dirty(offering_id, "basic_info", "课程基本信息已编辑")
    return {"status": "ok"}


@app.get("/api/offerings/{offering_id}/tasks")
def list_tasks(offering_id: int):
    return store.rows("SELECT * FROM tasks WHERE offering_id=? ORDER BY seq", (offering_id,))


@app.get("/api/offerings/{offering_id}/sessions")
def list_sessions(offering_id: int):
    return store.rows("SELECT * FROM sessions WHERE offering_id=? ORDER BY week_no, id", (offering_id,))


@app.get("/api/offerings/{offering_id}/curriculum-units")
def list_curriculum_units(offering_id: int):
    units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=? ORDER BY seq", (offering_id,))
    updates = store.rows(
        "SELECT id, topic, update_type, suggested_content, reason, confidence, status "
        "FROM content_updates WHERE offering_id=? AND status='待审核'",
        (offering_id,),
    )
    for u in units:
        warnings = []
        focus = (u.get("revised_focus") or "") + " " + (u.get("source_objectives") or "") + " " + (u.get("source_skills") or "") + " " + (u.get("project_title") or "")
        for upd in updates:
            topic = (upd.get("topic") or "").strip()
            if topic and topic in focus:
                warnings.append({
                    "id": upd["id"],
                    "topic": topic,
                    "update_type": upd["update_type"],
                    "suggested_content": (upd.get("suggested_content") or "")[:200],
                    "reason": (upd.get("reason") or "")[:150],
                    "confidence": upd["confidence"],
                })
        u["content_warnings"] = warnings
    return units


@app.get("/api/offerings/{offering_id}/documents")
def list_documents(offering_id: int):
    return store.rows("SELECT * FROM generated_documents WHERE offering_id=? ORDER BY id", (offering_id,))


@app.get("/api/offerings/{offering_id}/template-files")
def list_template_files(offering_id: int):
    return store.rows("SELECT * FROM template_files WHERE offering_id=? ORDER BY id", (offering_id,))


# ============================================================
# 模板库与契约管理（阶段1）
# ============================================================

@app.get("/api/template-library")
def list_template_library():
    rows = store.rows(
        """SELECT l.*,
               c.id AS contract_id, c.version AS contract_version, c.status AS contract_status,
               c.slot_count, c.parsed_at, c.confirmed_at,
               (SELECT COUNT(*) FROM contract_slots s WHERE s.contract_id=c.id AND s.approval_status='已确认') AS slot_confirmed
        FROM template_library l
        LEFT JOIN template_contracts c ON c.library_id=l.id
             AND c.version=(SELECT MAX(version) FROM template_contracts WHERE library_id=l.id)
        ORDER BY l.doc_type, l.version_label, l.id"""
    )
    for row in rows:
        row["exists"] = Path(row["file_path"]).exists()
    return rows


@app.post("/api/template-library/scan")
def scan_template_library(directory: str = ""):
    target = directory or str(TEMPLATE_DIR)
    return scan_templates(target)


class LibraryTemplateIn(BaseModel):
    file_path: str
    doc_type: str = "其他"
    name: str = ""
    version_label: str = ""
    notes: str = ""


@app.post("/api/template-library")
def add_library_template(data: LibraryTemplateIn):
    path = Path(data.file_path)
    if not path.is_file():
        raise AppError("文件不存在")
    if path.suffix.lower() not in (".docx", ".doc", ".xlsx"):
        raise AppError("仅支持 docx / doc / xlsx 模板")
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    existing = store.rows("SELECT id FROM template_library WHERE file_path=?", (str(path),))
    if existing:
        raise AppError("该文件已在模板库中")
    with store.connect() as db:
        cursor = db.execute(
            """INSERT INTO template_library (doc_type,name,version_label,file_path,file_hash,file_format,status,notes)
            VALUES (?,?,?,?,?,?,?,?)""",
            (
                data.doc_type, data.name or path.stem, data.version_label,
                str(path), digest, path.suffix.lower().lstrip("."), "待解析", data.notes,
            ),
        )
        db.commit()
        return {"id": cursor.lastrowid}


@app.put("/api/template-library/{library_id}")
def update_library_template(library_id: int, data: LibraryTemplateIn):
    row = store.rows("SELECT id FROM template_library WHERE id=?", (library_id,))
    if not row:
        raise NotFoundError("模板不存在")
    with store.connect() as db:
        db.execute(
            """UPDATE template_library SET doc_type=?,name=?,version_label=?,notes=?,updated_at=CURRENT_TIMESTAMP
            WHERE id=?""",
            (data.doc_type, data.name, data.version_label, data.notes, library_id),
        )
        db.commit()
    return {"ok": True}


@app.delete("/api/template-library/{library_id}")
def delete_library_template(library_id: int):
    row = store.rows("SELECT id FROM template_library WHERE id=?", (library_id,))
    if not row:
        raise NotFoundError("模板不存在")
    in_use = store.rows("SELECT id FROM template_files WHERE library_id=?", (library_id,))
    if in_use:
        raise AppError("该模板已被课程选用，请先解除选用关系")
    with store.connect() as db:
        db.execute("DELETE FROM template_library WHERE id=?", (library_id,))
        db.commit()
    return {"ok": True}


@app.post("/api/template-library/{library_id}/parse")
def parse_library_template(library_id: int):
    try:
        return parse_contract(library_id)
    except ValueError as exc:
        raise AppError(str(exc))


@app.get("/api/template-library/{library_id}/contract")
def get_library_contract(library_id: int):
    contracts = store.rows(
        "SELECT * FROM template_contracts WHERE library_id=? ORDER BY version DESC", (library_id,)
    )
    if not contracts:
        raise NotFoundError("该模板尚未解析")
    contract = dict(contracts[0])
    contract["structural_json"] = parse_json_field(contract.get("structural_json"), {})
    contract["content_json"] = parse_json_field(contract.get("content_json"), {})
    contract["slots"] = store.rows(
        "SELECT * FROM contract_slots WHERE contract_id=? ORDER BY sort_order, id", (contract["id"],)
    )
    for slot in contract["slots"]:
        slot["structure_json"] = parse_json_field(slot.get("structure_json"), {})
        slot["format_json"] = parse_json_field(slot.get("format_json"), {})
    return contract


@app.post("/api/template-contracts/{contract_id}/confirm")
def confirm_contract(contract_id: int):
    contracts = store.rows("SELECT * FROM template_contracts WHERE id=?", (contract_id,))
    if not contracts:
        raise NotFoundError("契约不存在")
    contract = contracts[0]
    unconfirmed = store.rows(
        "SELECT id FROM contract_slots WHERE contract_id=? AND approval_status='待确认' AND confidence='低'",
        (contract_id,),
    )
    if unconfirmed:
        raise AppError(f"存在 {len(unconfirmed)} 个低置信度槽位未确认，请先逐一核对")
    with store.connect() as db:
        db.execute(
            "UPDATE template_contracts SET status='已确认',confirmed_at=CURRENT_TIMESTAMP WHERE id=?",
            (contract_id,),
        )
        db.execute(
            "UPDATE template_contracts SET status='草稿' WHERE library_id=? AND id<>?",
            (contract["library_id"], contract_id),
        )
        db.commit()
    return {"ok": True}


class SlotOverrideIn(BaseModel):
    field_name: Optional[str] = None
    classification: Optional[str] = None
    content_req: Optional[str] = None
    format_json: Optional[dict] = None
    confidence: Optional[str] = None
    required: Optional[bool] = None
    approval_status: Optional[str] = None


@app.put("/api/contract-slots/{slot_id}")
def override_contract_slot(slot_id: int, data: SlotOverrideIn):
    slot = store.rows("SELECT * FROM contract_slots WHERE id=?", (slot_id,))
    if not slot:
        raise NotFoundError("槽位不存在")
    updates, params = [], []
    if data.field_name is not None:
        updates.append("field_name=?")
        params.append(data.field_name)
    if data.classification is not None:
        if data.classification not in ("A", "B", "C", "人工"):
            raise AppError("分类仅支持 A / B / C / 人工")
        updates.append("classification=?")
        params.append(data.classification)
    if data.content_req is not None:
        updates.append("content_req=?")
        params.append(data.content_req)
    if data.format_json is not None:
        updates.append("format_json=?")
        params.append(json.dumps(data.format_json, ensure_ascii=False))
    if data.confidence is not None:
        if data.confidence not in ("高", "中", "低"):
            raise AppError("置信度仅支持 高 / 中 / 低")
        updates.append("confidence=?")
        params.append(data.confidence)
    if data.required is not None:
        updates.append("required=?")
        params.append(1 if data.required else 0)
    if data.approval_status is not None:
        if data.approval_status not in ("待确认", "已确认", "需修改"):
            raise AppError("状态仅支持 待确认 / 已确认 / 需修改")
        updates.append("approval_status=?")
        params.append(data.approval_status)
    if not updates:
        return {"ok": True}
    updates.append("manual_override=1")
    params.append(slot_id)
    with store.connect() as db:
        db.execute(f"UPDATE contract_slots SET {', '.join(updates)} WHERE id=?", params)
        db.commit()
    return {"ok": True}


@app.get("/api/offerings/{offering_id}/source-files")
def list_source_files(offering_id: int):
    return store.rows("SELECT * FROM source_files WHERE offering_id=? ORDER BY id", (offering_id,))


@app.get("/api/offerings/{offering_id}/resources")
def list_offering_resources(offering_id: int):
    return store.rows("SELECT * FROM resource_items WHERE offering_id=? ORDER BY id", (offering_id,))


@app.get("/api/offerings/{offering_id}/content-model")
def get_content_model(offering_id: int):
    rows = store.rows("SELECT * FROM course_content_models WHERE offering_id=?", (offering_id,))
    if not rows:
        raise HTTPException(404, "尚未生成课程语义模型")
    model = dict(rows[0])
    model["model_json"] = parse_json_field(model.get("model_json"), {})
    return model


class ContentModelUpdate(BaseModel):
    ability_outcomes: List[str] = []
    teaching_methods: List[str] = []


@app.put("/api/offerings/{offering_id}/content-model")
def update_content_model(offering_id: int, body: ContentModelUpdate):
    _get_offering(offering_id)
    rows = store.rows("SELECT model_json FROM course_content_models WHERE offering_id=?", (offering_id,))
    if not rows:
        raise AppError("尚未生成课程内容模型，无法编辑。")
    model = parse_json_field(rows[0]["model_json"], {})
    if not isinstance(model, dict):
        model = {}
    model["ability_outcomes"] = [v.strip() for v in body.ability_outcomes if v.strip()]
    model["teaching_methods"] = [v.strip() for v in body.teaching_methods if v.strip()]
    with store.connect() as db:
        db.execute(
            "UPDATE course_content_models SET model_json=?, review_status='待检查' WHERE offering_id=?",
            (json.dumps(model, ensure_ascii=False), offering_id),
        )
        db.commit()
    return {"status": "ok"}


@app.get("/api/offerings/{offering_id}/drafts")
def list_drafts(offering_id: int, document_type: str = ""):
    sql = "SELECT * FROM authored_sections WHERE offering_id=?"
    params = [offering_id]
    if document_type:
        sql += " AND document_type=?"
        params.append(document_type)
    return store.rows(sql + " ORDER BY id", params)


@app.get("/api/offerings/{offering_id}/quality-issues")
def list_quality_issues(offering_id: int):
    return store.rows("SELECT * FROM quality_issues WHERE offering_id=? ORDER BY id DESC", (offering_id,))


# ============================================================
# 工作流操作
# ============================================================

def _get_offering(offering_id):
    """获取课程实例，不存在则抛404"""
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))
    if not offering:
        raise HTTPException(404, "课程不存在")
    return offering[0]


@app.post("/api/offerings/{offering_id}/rebuild-schedule")
def api_rebuild_schedule(offering_id: int):
    try:
        rebuild_schedule(_get_offering(offering_id))
        store.mark_dirty(offering_id, "schedule", "重新导入了排课数据")
        store.clear_dirty_by_action(offering_id, "rebuild_schedule")
        return {"status": "ok"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


@app.post("/api/offerings/{offering_id}/rebuild-resource-index")
def api_rebuild_resources(offering_id: int):
    try:
        build_resource_index(_get_offering(offering_id))
        store.mark_dirty(offering_id, "resources", "重新索引了教材资源")
        store.clear_dirty(offering_id, "resources")
        return {"status": "ok"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


@app.post("/api/offerings/{offering_id}/rebuild-curriculum-review")
def api_rebuild_review(offering_id: int):
    try:
        build_curriculum_review(_get_offering(offering_id))
        store.mark_dirty(offering_id, "review", "重建了蓝本审查")
        store.clear_dirty(offering_id, "review")
        return {"status": "ok"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


# 重建基础含AI内容生成（每任务一次在线调用），耗时远超前端超时，
# 改为后台线程执行 + 前端轮询进度，避免axios超时报错和重复触发
_foundation_jobs: dict = {}
_foundation_lock = threading.Lock()


def _run_foundation_job(offering_id: int):
    try:
        rebuild_generation_foundation(offering_id)
        store.clear_dirty(offering_id, "foundation")
        store.clear_dirty(offering_id, "templates")
        with _foundation_lock:
            _foundation_jobs[offering_id] = {"running": False, "error": ""}
    except Exception as e:
        with _foundation_lock:
            _foundation_jobs[offering_id] = {"running": False, "error": str(e)}


@app.post("/api/offerings/{offering_id}/rebuild-foundation")
def api_rebuild_foundation(offering_id: int):
    _get_offering(offering_id)
    with _foundation_lock:
        job = _foundation_jobs.get(offering_id)
        if job and job.get("running"):
            raise HTTPException(409, "重建生成基础正在后台执行，请等待完成。")
        _foundation_jobs[offering_id] = {"running": True, "error": ""}
    threading.Thread(target=_run_foundation_job, args=(offering_id,), daemon=True).start()
    return {"status": "started"}


@app.get("/api/offerings/{offering_id}/foundation-status")
def api_foundation_status(offering_id: int):
    _get_offering(offering_id)
    with _foundation_lock:
        job = _foundation_jobs.get(offering_id)
    running = bool(job and job.get("running"))
    stages = store.get_pipeline_status(offering_id)
    if not running:
        # 服务重启会让内存任务标记丢失，把卡在running的阶段标记为中断
        for stage, info in stages.items():
            if info.get("status") == "running":
                store.set_pipeline_stage(offering_id, stage, "failed", "服务重启导致中断，请重新重建。")
                info["status"] = "failed"
                info["error_message"] = "服务重启导致中断，请重新重建。"
    return {
        "running": running,
        "error": (job or {}).get("error", ""),
        "stages": stages,
    }


class GenerateDocumentsBody(BaseModel):
    document_types: Optional[List[str]] = None
    week_no: Optional[int] = None


@app.post("/api/offerings/{offering_id}/generate-documents")
def api_generate_documents(offering_id: int, body: Optional[GenerateDocumentsBody] = None):
    try:
        offering = _get_offering(offering_id)
        is_training = offering.get("offering_kind") == "实训课程"

        document_types = body.document_types if body and body.document_types else None
        week_no = body.week_no if body else None
        valid_types = {"课程标准", "授课计划", "教学设计"}
        if document_types is not None:
            unknown = [dt for dt in document_types if dt not in valid_types]
            if unknown:
                raise HTTPException(400, f"不支持的文档类型：{'、'.join(unknown)}")
            if not document_types:
                raise HTTPException(400, "未选择任何要生成的文档。")
        if week_no is not None and document_types is not None and "教学设计" not in document_types:
            raise HTTPException(400, "仅在选择生成教学设计时才能指定周次。")
        if week_no is not None and week_no < 1:
            raise HTTPException(400, "周次必须大于等于1。")

        if not is_training:
            tasks = store.rows("SELECT * FROM tasks WHERE offering_id=? ORDER BY seq", (offering_id,))
            sessions = store.rows("SELECT * FROM sessions WHERE offering_id=? ORDER BY week_no, id", (offering_id,))
            template_files = store.rows("SELECT * FROM template_files WHERE offering_id=?", (offering_id,))
            curriculum_units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=? ORDER BY seq", (offering_id,))
            blockers = generation_readiness(offering, tasks, sessions, template_files, curriculum_units, document_types=document_types)
            if blockers:
                raise HTTPException(400, "生成条件不满足: " + "; ".join(blockers))
            if week_no is not None:
                week_tasks = [t for t in tasks if t.get("week_no") == week_no]
                if not week_tasks:
                    known_weeks = sorted({t["week_no"] for t in tasks if t.get("week_no")})
                    raise HTTPException(400, f"第{week_no}周没有教学任务。当前有任务的周次：{'、'.join(map(str, known_weeks))}")

        names = generate_offering_documents(offering_id, document_types=document_types, week_no=week_no)
        store.clear_dirty(offering_id, "basic_info")
        store.clear_dirty(offering_id, "schedule")
        store.clear_dirty(offering_id, "teacher_name")
        return {"status": "ok", "paths": {k: str(v) for k, v in names.items()}}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


@app.get("/api/offerings/{offering_id}/generation-readiness")
def api_generation_readiness(offering_id: int, document_types: Optional[str] = None):
    try:
        offering = _get_offering(offering_id)
        if offering.get("offering_kind") == "实训课程":
            return {"blockers": [], "content_model": None, "templates": {}}
        selected = [dt for dt in document_types.split(",") if dt] if document_types else None
        tasks = store.rows("SELECT * FROM tasks WHERE offering_id=? ORDER BY seq", (offering_id,))
        sessions = store.rows("SELECT * FROM sessions WHERE offering_id=? ORDER BY week_no, id", (offering_id,))
        template_files = store.rows("SELECT * FROM template_files WHERE offering_id=?", (offering_id,))
        curriculum_units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=? ORDER BY seq", (offering_id,))
        blockers = generation_readiness(offering, tasks, sessions, template_files, curriculum_units, document_types=selected)
        model = store.rows("SELECT review_status FROM course_content_models WHERE offering_id=?", (offering_id,))
        templates = {}
        for row in template_files:
            analysis = store.rows(
                "SELECT analysis_status FROM template_analyses WHERE template_file_id=?",
                (row["id"],),
            )
            templates[row["document_type"]] = {
                "template_file_id": row["id"],
                "analysis_status": analysis[0]["analysis_status"] if analysis else "未分析",
            }
        return {
            "blockers": blockers,
            "content_model": {"review_status": model[0]["review_status"]} if model else None,
            "templates": templates,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


@app.post("/api/offerings/{offering_id}/confirm-content-model")
def api_confirm_content_model(offering_id: int):
    try:
        _get_offering(offering_id)
        model = store.rows("SELECT review_status FROM course_content_models WHERE offering_id=?", (offering_id,))
        if not model:
            raise HTTPException(400, "尚未生成课程内容模型，无法确认。")
        with store.connect() as db:
            db.execute(
                "UPDATE course_content_models SET review_status='已确认' WHERE offering_id=?",
                (offering_id,),
            )
            db.commit()
        return {"status": "ok"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


@app.post("/api/template-files/{template_file_id}/confirm-analysis")
def api_confirm_template_analysis(template_file_id: int):
    try:
        template = store.rows("SELECT id,document_type FROM template_files WHERE id=?", (template_file_id,))
        if not template:
            raise HTTPException(404, "模板文件不存在。")
        analysis = store.rows("SELECT template_file_id FROM template_analyses WHERE template_file_id=?", (template_file_id,))
        if not analysis:
            raise HTTPException(400, "该模板尚未完成规则分析，无法确认。")
        with store.connect() as db:
            db.execute(
                "UPDATE template_analyses SET analysis_status='已确认' WHERE template_file_id=?",
                (template_file_id,),
            )
            db.commit()
        return {"status": "ok"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


@app.post("/api/offerings/{offering_id}/build-tasks")
def api_build_tasks(offering_id: int, replace: bool = False):
    try:
        build_tasks(_get_offering(offering_id), replace=replace)
        store.clear_dirty(offering_id, "tasks")
        return {"status": "ok"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, str(e))


@app.post("/api/offerings/{offering_id}/reset-workflow")
def api_reset_workflow(offering_id: int):
    try:
        with store.connect() as db:
            db.execute("DELETE FROM curriculum_units WHERE offering_id=?", (offering_id,))
            db.execute("DELETE FROM tasks WHERE offering_id=?", (offering_id,))
            db.execute("DELETE FROM course_content_models WHERE offering_id=?", (offering_id,))
            db.execute("DELETE FROM authored_sections WHERE offering_id=?", (offering_id,))
            db.execute("DELETE FROM generated_documents WHERE offering_id=?", (offering_id,))
            db.execute("DELETE FROM template_rules WHERE template_file_id IN (SELECT id FROM template_files WHERE offering_id=?)", (offering_id,))
            db.execute("DELETE FROM template_slots WHERE template_file_id IN (SELECT id FROM template_files WHERE offering_id=?)", (offering_id,))
            db.execute("DELETE FROM dirty_flags WHERE offering_id=?", (offering_id,))
            db.commit()
        return {"status": "ok"}
    except Exception as e:
        raise HTTPException(400, str(e))


@app.post("/api/offerings/{offering_id}/approve-all-units")
def api_approve_units(offering_id: int):
    units = store.rows("SELECT id, suggested_hours FROM curriculum_units WHERE offering_id=? AND approval_status<>'已确认'", (offering_id,))
    if not units:
        return {"status": "ok", "approved": 0}
    total = sum(u["suggested_hours"] or 0 for u in units)
    offering_hours = store.rows("SELECT total_hours FROM offerings WHERE id=?", (offering_id,))[0]["total_hours"]
    if total != offering_hours:
        raise HTTPException(400, f"学时不一致: 单元合计{total}≠课程总学时{offering_hours}")
    with store.connect() as db:
        for u in units:
            db.execute("UPDATE curriculum_units SET approval_status='已确认' WHERE id=?", (u["id"],))
        db.commit()
    return {"status": "ok", "approved": len(units)}


class UnitUpdate(BaseModel):
    project_title: str = ""
    suggested_hours: int = 0
    source_objectives: str = ""
    source_skills: str = ""
    revised_focus: str = ""
    rationale: str = ""
    approval_status: str = ""


@app.put("/api/curriculum-units/{unit_id}")
def update_unit(unit_id: int, body: UnitUpdate):
    unit = store.rows("SELECT * FROM curriculum_units WHERE id=?", (unit_id,))
    if not unit:
        raise HTTPException(404, "蓝本单元不存在")
    fields = []
    values = []
    for name in ("project_title", "suggested_hours", "source_objectives", "source_skills", "revised_focus", "rationale", "approval_status"):
        val = getattr(body, name)
        if val or name == "suggested_hours":
            fields.append(f"{name}=?")
            values.append(val)
    if fields:
        values.append(unit_id)
        with store.connect() as db:
            db.execute(f"UPDATE curriculum_units SET {','.join(fields)} WHERE id=?", values)
            db.commit()
    store.mark_dirty(unit[0]["offering_id"], "review", "修改了蓝本单元")
    return {"status": "ok"}


# ============================================================
# 考勤
# ============================================================

@app.get("/api/attendance")
def get_attendance(offering_id: int, class_name: str = "", lesson_date: str = ""):
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))
    if not offering:
        raise HTTPException(404, "课程实例不存在")
    offering = offering[0]

    # 获取班级列表
    classes = [r["class_name"] for r in store.rows("SELECT DISTINCT class_name FROM students WHERE active=1 AND class_name<>'' ORDER BY class_name")]
    if not class_name and classes:
        class_name = classes[0]

    # 获取已确认课次
    sessions = store.rows(
        "SELECT lesson_date, week_no, periods, hours, classroom FROM sessions WHERE offering_id=? AND status='已确认' AND lesson_date<>'' ORDER BY lesson_date, id",
        (offering_id,),
    )

    # 获取学生
    students = store.rows("SELECT * FROM students WHERE class_name=? AND active=1 ORDER BY student_no, id", (class_name,)) if class_name else []

    # 获取已有考勤记录
    records = {}
    if lesson_date:
        for r in store.rows("SELECT * FROM attendance_records WHERE offering_id=? AND lesson_date=?", (offering_id, lesson_date)):
            records[r["student_id"]] = r

    # 考勤规则
    rules = store.rows("SELECT * FROM attendance_rules WHERE offering_id=? ORDER BY sort_order, id", (offering_id,))

    # 成绩构成
    components = store.rows("SELECT * FROM grade_components WHERE offering_id=? ORDER BY sort_order, id", (offering_id,))

    # 请假记录
    leaves = store.rows("SELECT sl.*, s.student_name, s.student_no FROM student_leave_periods sl JOIN students s ON s.id=sl.student_id WHERE sl.offering_id=? ORDER BY sl.start_date DESC", (offering_id,))

    return {
        "offering": offering,
        "classes": classes,
        "selected_class": class_name,
        "sessions": sessions,
        "lesson_date": lesson_date,
        "students": students,
        "records": records,
        "rules": rules,
        "components": components,
        "leaves": leaves,
    }


class AttendanceSave(BaseModel):
    student_ids: list[int]
    statuses: dict[str, str]
    notes: dict[str, str]
    scores: dict[str, float] = {}


@app.post("/api/attendance")
def save_attendance(offering_id: int, class_name: str, lesson_date: str, body: AttendanceSave):
    with store.connect() as db:
        for sid in body.student_ids:
            sid_str = str(sid)
            status = body.statuses.get(sid_str, "出勤")
            note = body.notes.get(sid_str, "")
            score = body.scores.get(sid_str, 0)
            db.execute(
                "INSERT INTO attendance_records (offering_id, student_id, lesson_date, status, score, notes) "
                "VALUES (?,?,?,?,?,?) ON CONFLICT(offering_id, student_id, lesson_date) "
                "DO UPDATE SET status=excluded.status, score=excluded.score, notes=excluded.notes, updated_at=CURRENT_TIMESTAMP",
                (offering_id, sid, lesson_date, status, score, note),
            )
        db.commit()
    return {"status": "ok"}


# ============================================================
# 学生
# ============================================================

@app.get("/api/students")
def list_students(class_name: str = ""):
    classes = [r["class_name"] for r in store.rows("SELECT DISTINCT class_name FROM students WHERE active=1 AND class_name<>'' ORDER BY class_name")]
    if not class_name and classes:
        class_name = classes[0]
    students = store.rows("SELECT * FROM students WHERE class_name=? AND active=1 ORDER BY student_no, id", (class_name,)) if class_name else []
    return {"classes": classes, "selected_class": class_name, "students": students}


class StudentAdd(BaseModel):
    class_name: str
    lines: str
    offering_id: int = 0


@app.post("/api/students")
def add_students(body: StudentAdd):
    count = 0
    errors = []
    with store.connect() as db:
        for line in body.lines.strip().splitlines():
            parts = line.strip().replace("\t", " ").split()
            if len(parts) < 2:
                continue
            student_no = parts[0]
            name = parts[1]
            gender = parts[2] if len(parts) > 2 else ""
            existing = store.rows("SELECT id FROM students WHERE class_name=? AND student_no=?", (body.class_name, student_no))
            try:
                if existing:
                    db.execute("UPDATE students SET student_name=?, gender=?, active=1 WHERE id=?", (name, gender, existing[0]["id"]))
                else:
                    db.execute("INSERT INTO students (class_name, student_no, student_name, gender) VALUES (?,?,?,?)", (body.class_name, student_no, name, gender))
                count += 1
            except Exception as e:
                errors.append(f"学号{student_no}: {e}")
        db.commit()
    return {"status": "ok", "added": count, "errors": errors}


class StudentUpdate(BaseModel):
    student_name: str
    gender: str = ""


@app.put("/api/students/{student_id}")
def update_student(student_id: int, body: StudentUpdate):
    with store.connect() as db:
        db.execute("UPDATE students SET student_name=?, gender=? WHERE id=?", (body.student_name, body.gender, student_id))
        db.commit()
    return {"status": "ok"}


@app.delete("/api/students/{student_id}")
def delete_student(student_id: int):
    with store.connect() as db:
        db.execute("UPDATE students SET active=0 WHERE id=?", (student_id,))
        db.commit()
    return {"status": "ok"}


# ============================================================
# 成绩分析
# ============================================================

@app.get("/api/grade-analysis")
def list_grade_analysis():
    return store.rows(
        "SELECT ga.*, o.term, o.course_name, o.major FROM grade_analysis_documents ga "
        "JOIN offerings o ON o.id=ga.offering_id ORDER BY ga.generated_at DESC, ga.id DESC"
    )


@app.post("/api/grade-analysis/generate")
async def generate_grade_analysis_api(
    offering_id: int = Form(...),
    grade_pdf: UploadFile = File(...),
    exam_date: str = Form(""),
    question_source: str = Form("自命题"),
    exam_mode: str = Form("其他方式"),
    marking_mode: str = Form("教师本人自阅"),
):
    tmp_path = await save_upload_file(grade_pdf, "grade", offering_id)
    try:
        result = generate_grade_analysis(offering_id, str(tmp_path), exam_date, question_source, exam_mode, marking_mode)
        return {"status": "ok", "result": result}
    except Exception as e:
        raise HTTPException(400, str(e))
    finally:
        tmp_path.unlink(missing_ok=True)


@app.delete("/api/grade-analysis/{document_id}")
def delete_grade_analysis(document_id: int):
    with store.connect() as db:
        db.execute("DELETE FROM grade_analysis_documents WHERE id=?", (document_id,))
        db.commit()
    return {"status": "ok"}


# ============================================================
# 实训资料
# ============================================================

@app.get("/api/training-materials")
def list_training_materials():
    return store.rows(
        "SELECT td.*, o.term, o.course_name, o.major FROM training_documents td "
        "JOIN offerings o ON o.id=td.offering_id "
        "WHERE COALESCE(o.offering_kind,'普通课程')='实训课程' "
        "ORDER BY td.generated_at DESC, td.id DESC"
    )


@app.get("/api/training-offerings")
def list_training_offerings():
    return store.rows(
        "SELECT * FROM offerings WHERE COALESCE(offering_kind,'普通课程')='实训课程' "
        "ORDER BY term DESC, course_name, major, id DESC"
    )


class TrainingGenerate(BaseModel):
    offering_id: int
    source_dir: str
    class_name: str = ""


@app.post("/api/training-materials/generate")
def generate_training_materials_api(body: TrainingGenerate):
    try:
        result = generate_training_materials(body.offering_id, body.source_dir, body.class_name)
        return {"status": "ok", "result": result}
    except Exception as e:
        raise HTTPException(400, str(e))


@app.delete("/api/training-materials/{document_id}")
def delete_training_material(document_id: int):
    with store.connect() as db:
        db.execute("DELETE FROM training_documents WHERE id=?", (document_id,))
        db.commit()
    return {"status": "ok"}


# ============================================================
# 听课记录
# ============================================================

@app.get("/api/listening-records")
def list_listening_records():
    docs = store.rows(
        "SELECT lr.*, o.term, o.course_name, s.class_name, s.lesson_date "
        "FROM listening_records lr "
        "JOIN offerings o ON o.id=lr.offering_id "
        "JOIN sessions s ON s.id=lr.session_id "
        "ORDER BY lr.generated_at DESC, lr.id DESC"
    )
    return docs


@app.get("/api/listening-sessions")
def list_listening_sessions():
    return store.rows(
        "SELECT s.id, s.lesson_date, s.class_name, o.course_name, o.term, o.major "
        "FROM sessions s JOIN offerings o ON o.id=s.offering_id "
        "WHERE s.status='已确认' AND s.lesson_date<>'' "
        "AND COALESCE(o.offering_kind,'普通课程')<>'实训课程' "
        "ORDER BY s.lesson_date DESC, o.course_name, s.class_name, s.id"
    )


@app.get("/api/listening-templates")
def list_listening_templates():
    return [str(p) for p in listening_templates()]


class ListeningGenerate(BaseModel):
    session_id: int
    template_path: str


@app.post("/api/listening-records/generate")
def generate_listening_record_api(body: ListeningGenerate):
    try:
        result = generate_listening_record(body.session_id, body.template_path)
        return {"status": "ok", "result": result}
    except Exception as e:
        raise HTTPException(400, str(e))


@app.delete("/api/listening-records/{record_id}")
def delete_listening_record(record_id: int):
    with store.connect() as db:
        db.execute("DELETE FROM listening_records WHERE id=?", (record_id,))
        db.commit()
    return {"status": "ok"}


# ============================================================
# 设置
# ============================================================

@app.get("/api/settings")
def get_settings():
    cache_count = store.rows("SELECT COUNT(*) AS count FROM model_cache")[0]["count"]
    enhanced_count = store.rows("SELECT COUNT(*) AS c FROM authored_sections WHERE authoring_status='增强AI生成'")[0]["c"]
    return {
        "ollama_url": store.get_setting("ollama_url", "http://127.0.0.1:11434"),
        "ollama_model": store.get_setting("ollama_model", "qwen3:8b"),
        "ollama_available": ollama_available(),
        "installed_models": installed_models(),
        "cache_count": cache_count,
        "enhanced_generation": store.get_setting("enhanced_generation", "0") == "1",
        "enhanced_count": enhanced_count,
        "ai_curriculum_review": store.get_setting("ai_curriculum_review", "0") == "1",
        "teacher_name": store.get_setting("teacher_name", "杜媛"),
        "output_root": store.get_setting("output_root", ""),
        "teaching_arrangement_path": store.get_setting("teaching_arrangement_path", ""),
    }


class SettingsUpdate(BaseModel):
    ollama_url: str = ""
    ollama_model: str = ""
    teacher_name: str = ""
    output_root: str = ""


@app.put("/api/settings")
def update_settings(body: SettingsUpdate):
    if body.ollama_url:
        store.set_setting("ollama_url", body.ollama_url.strip())
    if body.ollama_model:
        store.set_setting("ollama_model", body.ollama_model.strip())
    if body.teacher_name:
        old_name = store.get_setting("teacher_name", "")
        store.set_setting("teacher_name", body.teacher_name.strip())
        if old_name != body.teacher_name.strip():
            for offering in store.rows("SELECT id FROM offerings"):
                store.mark_dirty(offering["id"], "teacher_name", "全局教师姓名已修改")
    if body.output_root:
        store.set_setting("output_root", body.output_root.strip())
    return {"status": "ok"}


@app.post("/api/settings/toggle-enhanced")
def toggle_enhanced():
    current = store.get_setting("enhanced_generation", "0")
    store.set_setting("enhanced_generation", "0" if current == "1" else "1")
    return {"status": "ok", "enabled": current != "1"}


@app.post("/api/settings/toggle-ai-review")
def toggle_ai_review():
    current = store.get_setting("ai_curriculum_review", "0")
    store.set_setting("ai_curriculum_review", "0" if current == "1" else "1")
    return {"status": "ok", "enabled": current != "1"}


# ============================================================
# 蓝本审查规则
# ============================================================

@app.get("/api/review-rules")
def list_review_rules():
    return store.rows("SELECT * FROM curriculum_review_rules ORDER BY is_active DESC, course_name")


@app.post("/api/review-rules")
def create_review_rule(body: dict):
    with store.connect() as db:
        cursor = db.execute(
            """INSERT INTO curriculum_review_rules
            (course_name, rule_name, ppt_group_mode, project_pattern, outline_keyword,
             title_extraction, objective_keyword, skill_keywords, modernization_tags, is_active)
            VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (body.get("course_name", ""), body.get("rule_name", ""),
             body.get("ppt_group_mode", "each"), body.get("project_pattern", ""),
             body.get("outline_keyword", "教学大纲"), body.get("title_extraction", "first_slide"),
             body.get("objective_keyword", "学习目标"), body.get("skill_keywords", ""),
             body.get("modernization_tags", ""), int(body.get("is_active", True))),
        )
        db.commit()
        return {"id": cursor.lastrowid}


@app.put("/api/review-rules/{rule_id}")
def update_review_rule(rule_id: int, body: dict):
    with store.connect() as db:
        db.execute(
            """UPDATE curriculum_review_rules SET
            course_name=?, rule_name=?, ppt_group_mode=?, project_pattern=?, outline_keyword=?,
            title_extraction=?, objective_keyword=?, skill_keywords=?, modernization_tags=?, is_active=?
            WHERE id=?""",
            (body.get("course_name", ""), body.get("rule_name", ""),
             body.get("ppt_group_mode", "each"), body.get("project_pattern", ""),
             body.get("outline_keyword", "教学大纲"), body.get("title_extraction", "first_slide"),
             body.get("objective_keyword", "学习目标"), body.get("skill_keywords", ""),
             body.get("modernization_tags", ""), int(body.get("is_active", True)), rule_id),
        )
        db.commit()
    return {"status": "ok"}


@app.delete("/api/review-rules/{rule_id}")
def delete_review_rule(rule_id: int):
    with store.connect() as db:
        db.execute("DELETE FROM curriculum_review_rules WHERE id=?", (rule_id,))
        db.commit()
    return {"status": "ok"}


@app.post("/api/offerings/{offering_id}/ai-review")
def trigger_ai_review(offering_id: int):
    from ai_curriculum_review import ai_review_curriculum
    try:
        count, hours = ai_review_curriculum(offering_id)
        return {"status": "ok", "units": count, "hours": hours}
    except Exception as e:
        raise HTTPException(400, str(e))


# ============================================================
# 课程类型
# ============================================================

@app.get("/api/course-types")
def list_course_types():
    return store.rows("SELECT * FROM course_types ORDER BY sort_order, id")


class CourseTypeCreate(BaseModel):
    name: str
    sort_order: int = 0


@app.post("/api/course-types")
def create_course_type(body: CourseTypeCreate):
    with store.connect() as db:
        cursor = db.execute("INSERT INTO course_types (name, sort_order) VALUES (?,?)", (body.name, body.sort_order))
        db.commit()
        return {"id": cursor.lastrowid}


class CourseTypeUpdate(BaseModel):
    name: str
    sort_order: int = 0


@app.put("/api/course-types/{type_id}")
def update_course_type(type_id: int, body: CourseTypeUpdate):
    with store.connect() as db:
        db.execute("UPDATE course_types SET name=?, sort_order=? WHERE id=?", (body.name, body.sort_order, type_id))
        db.commit()
    return {"status": "ok"}


@app.delete("/api/course-types/{type_id}")
def delete_course_type(type_id: int):
    with store.connect() as db:
        db.execute("DELETE FROM course_types WHERE id=?", (type_id,))
        db.commit()
    return {"status": "ok"}


# ============================================================
# 资源中心
# ============================================================

@app.get("/api/resources")
def list_resources(course: int = 0, kind: str = ""):
    sql = "SELECT ri.*, o.term, o.course_name FROM resource_items ri JOIN offerings o ON o.id=ri.offering_id WHERE 1=1"
    params = []
    if course:
        sql += " AND ri.offering_id=?"
        params.append(course)
    if kind:
        sql += " AND ri.resource_type=?"
        params.append(kind)
    sql += " ORDER BY o.term DESC, o.course_name, ri.id DESC"
    return store.rows(sql, params)


@app.get("/api/resource-types")
def list_resource_types():
    return [r["resource_type"] for r in store.rows("SELECT DISTINCT resource_type FROM resource_items WHERE resource_type<>'' ORDER BY resource_type")]


# ============================================================
# 作业成绩
# ============================================================

@app.get("/api/assignments/{assignment_id}")
def get_assignment(assignment_id: int, class_name: str = ""):
    rows = store.rows(
        "SELECT a.*, o.course_name, o.term, o.teaching_class FROM assignments a "
        "JOIN offerings o ON o.id=a.offering_id WHERE a.id=?",
        (assignment_id,),
    )
    if not rows:
        raise HTTPException(404, "作业不存在")
    assignment = rows[0]
    classes = [r["class_name"] for r in store.rows("SELECT DISTINCT class_name FROM students WHERE active=1 AND class_name<>'' ORDER BY class_name")]
    if not class_name and classes:
        class_name = classes[0]
    students = store.rows("SELECT * FROM students WHERE class_name=? AND active=1 ORDER BY student_no, id", (class_name,)) if class_name else []
    scores = {r["student_id"]: r for r in store.rows("SELECT * FROM assignment_scores WHERE assignment_id=?", (assignment_id,))}
    issues = store.rows("SELECT * FROM assignment_import_issues WHERE assignment_id=? ORDER BY id", (assignment_id,))
    return {"assignment": assignment, "classes": classes, "selected_class": class_name, "students": students, "scores": scores, "issues": issues}


class AssignmentScoresSave(BaseModel):
    scores: dict[str, float]


@app.post("/api/assignments/{assignment_id}/scores")
def save_assignment_scores(assignment_id: int, class_name: str, body: AssignmentScoresSave):
    with store.connect() as db:
        for sid_str, score in body.scores.items():
            sid = int(sid_str)
            db.execute(
                "INSERT INTO assignment_scores (assignment_id, student_id, score, source) VALUES (?,?,?,'手工录入') "
                "ON CONFLICT(assignment_id, student_id) DO UPDATE SET score=excluded.score, source='手工录入', updated_at=CURRENT_TIMESTAMP",
                (assignment_id, sid, score),
            )
        db.commit()
    return {"status": "ok"}


# ============================================================
# 批量操作
# ============================================================

@app.post("/api/offerings/{offering_id}/import-arrangement")
async def import_arrangement(offering_id: int, file: UploadFile = File(...)):
    tmp = await save_upload_file(file, "arr", offering_id)
    try:
        import_teaching_arrangement(_get_offering(offering_id), str(tmp))
        store.mark_dirty(offering_id, "schedule", "导入了教学安排表")
        store.mark_dirty(offering_id, "teacher_name", "教学安排表可能更新了教师姓名")
        return {"status": "ok"}
    except Exception as e:
        raise HTTPException(400, str(e))
    finally:
        tmp.unlink(missing_ok=True)


@app.post("/api/offerings/{offering_id}/import-progress")
async def import_progress(offering_id: int, file: UploadFile = File(...)):
    tmp = await save_upload_file(file, "prog", offering_id)
    try:
        import_progress_table(_get_offering(offering_id), str(tmp))
        store.mark_dirty(offering_id, "schedule", "导入了学期进程表")
        return {"status": "ok"}
    except Exception as e:
        raise HTTPException(400, str(e))
    finally:
        tmp.unlink(missing_ok=True)


@app.post("/api/offerings/{offering_id}/import-calendar")
async def import_calendar(offering_id: int, file: UploadFile = File(...)):
    tmp = await save_upload_file(file, "cal", offering_id)
    try:
        offering = _get_offering(offering_id)
        sessions = store.rows(
            "SELECT lesson_date FROM sessions WHERE offering_id=?",
            (offering_id,),
        )
        if sessions and not any(row["lesson_date"] for row in sessions):
            import_calendar_week_dates(offering, str(tmp))
        import_school_calendar(offering, str(tmp))
        store.mark_dirty(offering_id, "schedule", "导入了校历事件")
        return {"status": "ok"}
    except Exception as e:
        raise HTTPException(400, str(e))
    finally:
        tmp.unlink(missing_ok=True)


# ============================================================
# 打开文件夹
# ============================================================

class OpenLocation(BaseModel):
    offering_id: int
    kind: str
    document_id: int = 0


@app.post("/api/open-location")
def open_location(body: OpenLocation):
    output_root = Path(store.get_setting("output_root", str(Path(__file__).parent / "生成结果"))).resolve()

    if body.kind == "grade_analysis":
        rows = store.rows("SELECT output_path FROM grade_analysis_documents WHERE id=?", (body.document_id,))
    elif body.kind == "listening_record":
        rows = store.rows("SELECT output_path FROM listening_records WHERE id=?", (body.document_id,))
    elif body.kind == "training_material":
        rows = store.rows("SELECT output_path FROM training_documents WHERE id=?", (body.document_id,))
    else:
        offering = store.rows("SELECT * FROM offerings WHERE id=?", (body.offering_id,))
        if not offering:
            raise HTTPException(404, "课程不存在")
        path = (output_root / f"offering_{body.offering_id}").resolve()
        if not str(path).startswith(str(output_root)):
            raise HTTPException(403, "路径不在允许范围内")
        if path.exists():
            subprocess.Popen(["explorer", str(path)])
        return {"status": "ok"}

    if not rows:
        raise HTTPException(404, "文件不存在")
    target = Path(rows[0]["output_path"]).parent.resolve()
    if not str(target).startswith(str(output_root)):
        raise HTTPException(403, "路径不在允许范围内")
    if target.exists():
        subprocess.Popen(["explorer", str(target)])
    return {"status": "ok"}


# ============================================================
# 模板规则
# ============================================================

@app.get("/api/template-files/{template_file_id}/rules")
def list_template_rules(template_file_id: int):
    return store.rows("SELECT * FROM template_rules WHERE template_file_id=? ORDER BY seq", (template_file_id,))


@app.get("/api/template-files/{template_file_id}/slots")
def list_template_slots(template_file_id: int):
    return store.rows("SELECT * FROM template_slots WHERE template_file_id=? ORDER BY slot_key", (template_file_id,))


# ============================================================
# 任务 CRUD
# ============================================================

class TaskUpdate(BaseModel):
    chapter: str = ""
    title: str = ""
    hours: int = 0
    theory_hours: int = 0
    practice_hours: int = 0
    knowledge_goal: str = ""
    ability_goal: str = ""
    ideological_goal: str = ""
    quality_goal: str = ""


@app.put("/api/tasks/{task_id}")
def update_task(task_id: int, body: TaskUpdate):
    with store.connect() as db:
        db.execute(
            "UPDATE tasks SET chapter=?, title=?, hours=?, theory_hours=?, practice_hours=?, "
            "knowledge_goal=?, ability_goal=?, ideological_goal=?, quality_goal=? WHERE id=?",
            (body.chapter, body.title, body.hours, body.theory_hours, body.practice_hours,
             body.knowledge_goal, body.ability_goal, body.ideological_goal, body.quality_goal, task_id),
        )
        db.commit()
    task = store.rows("SELECT offering_id FROM tasks WHERE id=?", (task_id,))
    if task:
        store.mark_dirty(task[0]["offering_id"], "tasks", "修改了教学任务")
    return {"status": "ok"}


@app.delete("/api/tasks/{task_id}")
def delete_task(task_id: int):
    task = store.rows("SELECT offering_id FROM tasks WHERE id=?", (task_id,))
    with store.connect() as db:
        db.execute("DELETE FROM tasks WHERE id=?", (task_id,))
        db.commit()
    if task:
        store.mark_dirty(task[0]["offering_id"], "tasks", "删除了教学任务")
    return {"status": "ok"}


# ============================================================
# 源文件/模板文件管理
# ============================================================

class SourceFileCreate(BaseModel):
    offering_id: int
    source_type: str
    source_path: str
    required: bool = False
    notes: str = ""


@app.post("/api/source-files")
def create_source_file(body: SourceFileCreate):
    with store.connect() as db:
        cursor = db.execute(
            "INSERT INTO source_files (offering_id, source_type, source_path, required, notes) VALUES (?,?,?,?,?)",
            (body.offering_id, body.source_type, body.source_path, int(body.required), body.notes),
        )
        db.commit()
        return {"id": cursor.lastrowid}


class SourceFileUpdate(BaseModel):
    source_type: str
    source_path: str
    required: bool = False
    notes: str = ""


@app.put("/api/source-files/{file_id}")
def update_source_file(file_id: int, body: SourceFileUpdate):
    with store.connect() as db:
        db.execute("UPDATE source_files SET source_type=?, source_path=?, required=?, notes=? WHERE id=?", (body.source_type, body.source_path, int(body.required), body.notes, file_id))
        db.commit()
    file = store.rows("SELECT offering_id FROM source_files WHERE id=?", (file_id,))
    if file:
        store.mark_dirty(file[0]["offering_id"], "resources", "修改了源文件")
    return {"status": "ok"}


@app.delete("/api/source-files/{file_id}")
def delete_source_file(file_id: int):
    file = store.rows("SELECT offering_id FROM source_files WHERE id=?", (file_id,))
    with store.connect() as db:
        db.execute("DELETE FROM source_files WHERE id=?", (file_id,))
        db.commit()
    if file:
        store.mark_dirty(file[0]["offering_id"], "resources", "删除了源文件")
    return {"status": "ok"}


class TemplateFileCreate(BaseModel):
    offering_id: int
    document_type: str
    template_name: str
    template_path: str
    required: bool = False
    notes: str = ""


@app.post("/api/template-files")
def create_template_file(body: TemplateFileCreate):
    with store.connect() as db:
        cursor = db.execute(
            "INSERT INTO template_files (offering_id, document_type, template_name, template_path, required, notes) VALUES (?,?,?,?,?,?)",
            (body.offering_id, body.document_type, body.template_name, body.template_path, int(body.required), body.notes),
        )
        db.commit()
        return {"id": cursor.lastrowid}


class TemplateFileUpdate(BaseModel):
    document_type: str
    template_name: str
    template_path: str
    required: bool = False
    notes: str = ""


@app.put("/api/template-files/{file_id}")
def update_template_file(file_id: int, body: TemplateFileUpdate):
    with store.connect() as db:
        db.execute("UPDATE template_files SET document_type=?, template_name=?, template_path=?, required=?, notes=? WHERE id=?", (body.document_type, body.template_name, body.template_path, int(body.required), body.notes, file_id))
        db.commit()
    file = store.rows("SELECT offering_id FROM template_files WHERE id=?", (file_id,))
    if file:
        store.mark_dirty(file[0]["offering_id"], "templates", "修改了模板文件")
    return {"status": "ok"}


@app.delete("/api/template-files/{file_id}")
def delete_template_file(file_id: int):
    file = store.rows("SELECT offering_id FROM template_files WHERE id=?", (file_id,))
    with store.connect() as db:
        db.execute("DELETE FROM template_files WHERE id=?", (file_id,))
        db.commit()
    if file:
        store.mark_dirty(file[0]["offering_id"], "templates", "删除了模板文件")
    return {"status": "ok"}


# ============================================================
# 考勤规则/成绩构成
# ============================================================

class AttendanceRuleCreate(BaseModel):
    offering_id: int
    status: str
    deduction: float = 0
    sort_order: int = 0


@app.post("/api/attendance-rules")
def create_attendance_rule(body: AttendanceRuleCreate):
    with store.connect() as db:
        cursor = db.execute("INSERT INTO attendance_rules (offering_id, status, deduction, sort_order) VALUES (?,?,?,?)", (body.offering_id, body.status, body.deduction, body.sort_order))
        db.commit()
        return {"id": cursor.lastrowid}


@app.delete("/api/attendance-rules/{rule_id}")
def delete_attendance_rule(rule_id: int):
    with store.connect() as db:
        db.execute("DELETE FROM attendance_rules WHERE id=?", (rule_id,))
        db.commit()
    return {"status": "ok"}


class GradeComponentCreate(BaseModel):
    offering_id: int
    component_name: str
    weight: float = 0
    source_type: str = ""
    sort_order: int = 0


@app.post("/api/grade-components")
def create_grade_component(body: GradeComponentCreate):
    with store.connect() as db:
        cursor = db.execute("INSERT INTO grade_components (offering_id, component_name, weight, source_type, sort_order) VALUES (?,?,?,?,?)", (body.offering_id, body.component_name, body.weight, body.source_type, body.sort_order))
        db.commit()
        return {"id": cursor.lastrowid}


@app.delete("/api/grade-components/{component_id}")
def delete_grade_component(component_id: int):
    with store.connect() as db:
        db.execute("DELETE FROM grade_components WHERE id=?", (component_id,))
        db.commit()
    return {"status": "ok"}


# ============================================================
# 请假管理
# ============================================================

class LeaveCreate(BaseModel):
    offering_id: int
    student_id: int
    start_date: str
    end_date: str
    reason: str = ""


@app.post("/api/leaves")
def create_leave(body: LeaveCreate):
    with store.connect() as db:
        cursor = db.execute("INSERT INTO student_leave_periods (offering_id, student_id, start_date, end_date, reason) VALUES (?,?,?,?,?)", (body.offering_id, body.student_id, body.start_date, body.end_date, body.reason))
        db.commit()
        return {"id": cursor.lastrowid}


class LeaveUpdate(BaseModel):
    start_date: str
    end_date: str
    reason: str = ""


@app.put("/api/leaves/{leave_id}")
def update_leave(leave_id: int, body: LeaveUpdate):
    with store.connect() as db:
        db.execute("UPDATE student_leave_periods SET start_date=?, end_date=?, reason=? WHERE id=?", (body.start_date, body.end_date, body.reason, leave_id))
        db.commit()
    return {"status": "ok"}


@app.delete("/api/leaves/{leave_id}")
def delete_leave(leave_id: int):
    with store.connect() as db:
        db.execute("DELETE FROM student_leave_periods WHERE id=?", (leave_id,))
        db.commit()
    return {"status": "ok"}


# ============================================================
# 排课
# ============================================================

class SessionUpdate(BaseModel):
    week_no: int = 0
    lesson_date: str = ""
    classroom: str = ""
    status: str = ""
    session_type: str = ""
    periods: str = ""
    hours: int = 0


WEEKDAY_NAMES = ("周一", "周二", "周三", "周四", "周五", "周六", "周日")
SESSION_TYPES = ("正常排课", "补课", "调课", "停课")
SESSION_STATUSES = ("待确认", "已确认", "已取消")


@app.put("/api/sessions/{session_id}")
def update_session(session_id: int, body: SessionUpdate):
    if body.session_type and body.session_type not in SESSION_TYPES:
        raise AppError(f"排课类型不合法，可选：{'、'.join(SESSION_TYPES)}")
    if body.status and body.status not in SESSION_STATUSES:
        raise AppError(f"状态不合法，可选：{'、'.join(SESSION_STATUSES)}")
    if body.hours < 0 or body.week_no < 0:
        raise AppError("学时与周次不能为负数")
    weekday = ""
    if body.lesson_date:
        try:
            weekday = WEEKDAY_NAMES[date.fromisoformat(body.lesson_date).weekday()]
        except ValueError:
            raise AppError("上课日期格式应为 YYYY-MM-DD")
    fields = ["week_no", "lesson_date", "weekday", "classroom", "status", "session_type", "periods", "hours"]
    data = body.model_dump()
    data["weekday"] = weekday
    with store.connect() as db:
        row = db.execute("SELECT offering_id FROM sessions WHERE id=?", (session_id,)).fetchone()
        if row is None:
            raise NotFoundError("排课记录不存在")
        db.execute(
            f"UPDATE sessions SET {','.join(f + '=?' for f in fields)} WHERE id=?",
            [data[f] for f in fields] + [session_id],
        )
        db.commit()
        offering_id = row["offering_id"]
    store.mark_dirty(offering_id, "schedule", "排课数据已编辑")
    return {"status": "ok"}


# ============================================================
# 教学单元
# ============================================================

class CurriculumUnitUpdate(BaseModel):
    project_title: str = ""
    suggested_hours: int = 0
    review_action: str = ""
    revised_focus: str = ""
    rationale: str = ""


@app.put("/api/curriculum-units/{unit_id}")
def update_curriculum_unit(unit_id: int, body: CurriculumUnitUpdate):
    with store.connect() as db:
        db.execute(
            "UPDATE curriculum_units SET project_title=?, suggested_hours=?, review_action=?, revised_focus=?, rationale=? WHERE id=?",
            (body.project_title, body.suggested_hours, body.review_action, body.revised_focus, body.rationale, unit_id),
        )
        db.commit()
    return {"status": "ok"}


# ============================================================
# 内容更新建议
# ============================================================

class ContentUpdateReview(BaseModel):
    status: str = "已采纳"
    reviewed_by: str = ""


class ContentUpdateCreate(BaseModel):
    update_type: str = "内容更新"
    topic: str = ""
    original_summary: str = ""
    suggested_content: str = ""
    reason: str = ""
    source_urls: list = []
    confidence: float = 0.5
    related_chapters: list = []


@app.get("/api/offerings/{offering_id}/content-updates")
def list_content_updates_api(offering_id: int, status: Optional[str] = None):
    """获取课程的内容更新建议列表"""
    updates = store.list_content_updates(offering_id, status=status)
    # 解析 JSON 字段
    for u in updates:
        for field in ("source_urls", "related_chapters"):
            if isinstance(u.get(field), str):
                try:
                    u[field] = json.loads(u[field])
                except (json.JSONDecodeError, TypeError):
                    u[field] = []
    return {"items": updates, "total": len(updates)}


@app.post("/api/offerings/{offering_id}/content-updates/analyze")
def analyze_content_updates_api(offering_id: int):
    """触发内容更新分析（AI识别过时内容）"""
    count = analyze_content_updates(offering_id)
    return {"analyzed": True, "new_suggestions": count}


@app.post("/api/offerings/{offering_id}/content-updates")
def create_content_update_api(offering_id: int, body: ContentUpdateCreate):
    """手动添加一条内容更新建议"""
    update_id = store.add_content_update(
        offering_id=offering_id,
        update_type=body.update_type,
        topic=body.topic,
        original_summary=body.original_summary,
        suggested_content=body.suggested_content,
        reason=body.reason,
        source_urls=body.source_urls,
        confidence=body.confidence,
        related_chapters=body.related_chapters,
    )
    store.mark_dirty(offering_id, "foundation", "手动添加内容更新建议")
    return {"id": update_id}


@app.put("/api/content-updates/{update_id}/status")
def review_content_update_api(update_id: int, body: ContentUpdateReview):
    """审核更新建议：已采纳/已忽略/待审核"""
    if body.status not in ("已采纳", "已忽略", "待审核"):
        raise HTTPException(400, "无效的状态值")
    update = store.rows("SELECT * FROM content_updates WHERE id=?", (update_id,))
    if not update:
        raise HTTPException(404, "更新建议不存在")
    store.update_content_update_status(update_id, body.status, body.reviewed_by)
    # 如果状态变化涉及采纳/取消，标记文档需要重新生成
    if body.status in ("已采纳", "已忽略"):
        store.mark_dirty(update[0]["offering_id"], "foundation", f"内容更新建议{body.status}：{update[0]['topic']}")
    return {"ok": True}


@app.delete("/api/content-updates/{update_id}")
def delete_content_update_api(update_id: int):
    """删除一条更新建议"""
    update = store.rows("SELECT * FROM content_updates WHERE id=?", (update_id,))
    if not update:
        raise HTTPException(404, "更新建议不存在")
    store.delete_content_update(update_id)
    return {"ok": True}


# ============================================================
# 静态文件服务（Vue 前端构建产物）
# ============================================================

frontend_dist = Path(__file__).parent / "frontend" / "dist"
if frontend_dist.exists():
    app.mount("/", StaticFiles(directory=str(frontend_dist), html=True), name="frontend")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=HTTP_HOST, port=HTTP_PORT)

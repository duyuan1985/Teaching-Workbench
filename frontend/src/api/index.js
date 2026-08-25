import axios from 'axios'
import { ElMessage } from 'element-plus'

const http = axios.create({
  baseURL: '/api',
  timeout: 120000,
})

http.interceptors.response.use(
  (response) => response.data,
  (error) => {
    const msg = error.response?.data?.detail || error.message || '请求失败'
    ElMessage.error(msg)
    return Promise.reject(error)
  }
)

// ============================================================
// 课程实例
// ============================================================

export const offeringsApi = {
  list: (params) => http.get('/offerings', { params }),
  stats: () => http.get('/offerings/stats'),
  get: (id) => http.get(`/offerings/${id}`),
  create: (data) => http.post('/offerings', data),
  update: (id, data) => http.put(`/offerings/${id}`, data),
  terms: () => http.get('/terms'),
  tasks: (id) => http.get(`/offerings/${id}/tasks`),
  sessions: (id) => http.get(`/offerings/${id}/sessions`),
  units: (id) => http.get(`/offerings/${id}/curriculum-units`),
  documents: (id) => http.get(`/offerings/${id}/documents`),
  templateFiles: (id) => http.get(`/offerings/${id}/template-files`),
  sourceFiles: (id) => http.get(`/offerings/${id}/source-files`),
  resources: (id) => http.get(`/offerings/${id}/resources`),
  contentModel: (id) => http.get(`/offerings/${id}/content-model`),
  updateContentModel: (id, data) => http.put(`/offerings/${id}/content-model`, data),
  drafts: (id, docType) => http.get(`/offerings/${id}/drafts`, { params: { document_type: docType } }),
  qualityIssues: (id) => http.get(`/offerings/${id}/quality-issues`),
  rebuildSchedule: (id) => http.post(`/offerings/${id}/rebuild-schedule`),
  rebuildResources: (id) => http.post(`/offerings/${id}/rebuild-resource-index`),
  rebuildReview: (id) => http.post(`/offerings/${id}/rebuild-curriculum-review`),
  rebuildFoundation: (id) => http.post(`/offerings/${id}/rebuild-foundation`),
  foundationStatus: (id) => http.get(`/offerings/${id}/foundation-status`),
  generateDocuments: (id, data) => http.post(`/offerings/${id}/generate-documents`, data),
  generationReadiness: (id, documentTypes) => http.get(`/offerings/${id}/generation-readiness`, {
    params: documentTypes && documentTypes.length ? { document_types: documentTypes.join(',') } : {},
  }),
  confirmContentModel: (id) => http.post(`/offerings/${id}/confirm-content-model`),
  confirmTemplateAnalysis: (templateFileId) => http.post(`/template-files/${templateFileId}/confirm-analysis`),
  buildTasks: (id) => http.post(`/offerings/${id}/build-tasks?replace=true`),
  resetWorkflow: (id) => http.post(`/offerings/${id}/reset-workflow`),
  approveUnits: (id) => http.post(`/offerings/${id}/approve-all-units`),
  dirtyFlags: (id) => http.get(`/offerings/${id}/dirty-flags`),
  contentUpdates: (id, status) => http.get(`/offerings/${id}/content-updates`, { params: { status } }),
  analyzeContentUpdates: (id) => http.post(`/offerings/${id}/content-updates/analyze`),
  reviewContentUpdate: (updateId, status) => http.put(`/content-updates/${updateId}/status`, { status }),
  deleteContentUpdate: (updateId) => http.delete(`/content-updates/${updateId}`),
  updateUnit: (unitId, data) => http.put(`/curriculum-units/${unitId}`, data),
  importArrangement: (id, file) => {
    const form = new FormData()
    form.append('file', file)
    return http.post(`/offerings/${id}/import-arrangement`, form)
  },
  importProgress: (id, file) => {
    const form = new FormData()
    form.append('file', file)
    return http.post(`/offerings/${id}/import-progress`, form)
  },
  importCalendar: (id, file) => {
    const form = new FormData()
    form.append('file', file)
    return http.post(`/offerings/${id}/import-calendar`, form)
  },
}

// ============================================================
// 考勤
// ============================================================

export const attendanceApi = {
  get: (offeringId, className, lessonDate) => http.get('/attendance', { params: { offering_id: offeringId, class_name: className, lesson_date: lessonDate } }),
  save: (offeringId, className, lessonDate, data) => http.post('/attendance', data, { params: { offering_id: offeringId, class_name: className, lesson_date: lessonDate } }),
}

export const attendanceRulesApi = {
  create: (data) => http.post('/attendance-rules', data),
  delete: (id) => http.delete(`/attendance-rules/${id}`),
}

export const gradeComponentsApi = {
  create: (data) => http.post('/grade-components', data),
  delete: (id) => http.delete(`/grade-components/${id}`),
}

export const leavesApi = {
  create: (data) => http.post('/leaves', data),
  update: (id, data) => http.put(`/leaves/${id}`, data),
  delete: (id) => http.delete(`/leaves/${id}`),
}

// ============================================================
// 学生
// ============================================================

export const studentsApi = {
  list: (className) => http.get('/students', { params: { class_name: className } }),
  create: (data) => http.post('/students', data),
  update: (id, data) => http.put(`/students/${id}`, data),
  delete: (id) => http.delete(`/students/${id}`),
}

// ============================================================
// 成绩分析
// ============================================================

export const gradeAnalysisApi = {
  list: () => http.get('/grade-analysis'),
  generate: (formData) => http.post('/grade-analysis/generate', formData),
  delete: (id) => http.delete(`/grade-analysis/${id}`),
}

// ============================================================
// 实训资料
// ============================================================

export const trainingApi = {
  list: () => http.get('/training-materials'),
  offerings: () => http.get('/training-offerings'),
  generate: (data) => http.post('/training-materials/generate', data),
  delete: (id) => http.delete(`/training-materials/${id}`),
}

// ============================================================
// 听课记录
// ============================================================

export const listeningApi = {
  list: () => http.get('/listening-records'),
  sessions: () => http.get('/listening-sessions'),
  templates: () => http.get('/listening-templates'),
  generate: (data) => http.post('/listening-records/generate', data),
  delete: (id) => http.delete(`/listening-records/${id}`),
}

// ============================================================
// 设置
// ============================================================

export const settingsApi = {
  get: () => http.get('/settings'),
  update: (data) => http.put('/settings', data),
  toggleEnhanced: () => http.post('/settings/toggle-enhanced'),
}

// ============================================================
// 课程类型
// ============================================================

export const courseTypesApi = {
  list: () => http.get('/course-types'),
  create: (data) => http.post('/course-types', data),
  update: (id, data) => http.put(`/course-types/${id}`, data),
  delete: (id) => http.delete(`/course-types/${id}`),
}

// ============================================================
// 蓝本审查规则
// ============================================================

export const reviewRulesApi = {
  list: () => http.get('/review-rules'),
  create: (data) => http.post('/review-rules', data),
  update: (id, data) => http.put(`/review-rules/${id}`, data),
  delete: (id) => http.delete(`/review-rules/${id}`),
}

export const aiReviewApi = {
  toggle: () => http.post('/settings/toggle-ai-review'),
  trigger: (offeringId) => http.post(`/offerings/${offeringId}/ai-review`),
}

// ============================================================
// 资源
// ============================================================

export const resourcesApi = {
  list: (course, kind) => http.get('/resources', { params: { course, kind } }),
  types: () => http.get('/resource-types'),
}

// ============================================================
// 任务
// ============================================================

export const tasksApi = {
  update: (id, data) => http.put(`/tasks/${id}`, data),
  delete: (id) => http.delete(`/tasks/${id}`),
}

// ============================================================
// 会话/排课
// ============================================================

export const sessionsApi = {
  update: (id, data) => http.put(`/sessions/${id}`, data),
}

// ============================================================
// 教学单元
// ============================================================

export const unitsApi = {
  update: (id, data) => http.put(`/curriculum-units/${id}`, data),
}

// ============================================================
// 源文件/模板文件
// ============================================================

export const sourceFilesApi = {
  create: (data) => http.post('/source-files', data),
  update: (id, data) => http.put(`/source-files/${id}`, data),
  delete: (id) => http.delete(`/source-files/${id}`),
}

export const templateFilesApi = {
  create: (data) => http.post('/template-files', data),
  update: (id, data) => http.put(`/template-files/${id}`, data),
  delete: (id) => http.delete(`/template-files/${id}`),
  rules: (id) => http.get(`/template-files/${id}/rules`),
  slots: (id) => http.get(`/template-files/${id}/slots`),
}

// ============================================================
// 模板库与契约管理（阶段1）
// ============================================================

export const templateLibraryApi = {
  list: () => http.get('/template-library'),
  scan: (directory) => http.post('/template-library/scan', null, { params: { directory } }),
  create: (data) => http.post('/template-library', data),
  update: (id, data) => http.put(`/template-library/${id}`, data),
  delete: (id) => http.delete(`/template-library/${id}`),
  parse: (id) => http.post(`/template-library/${id}/parse`),
  contract: (id) => http.get(`/template-library/${id}/contract`),
  confirmContract: (contractId) => http.post(`/template-contracts/${contractId}/confirm`),
  updateSlot: (slotId, data) => http.put(`/contract-slots/${slotId}`, data),
}

// ============================================================
// 作业
// ============================================================

export const assignmentsApi = {
  get: (id, className) => http.get(`/assignments/${id}`, { params: { class_name: className } }),
  saveScores: (id, className, data) => http.post(`/assignments/${id}/scores`, data, { params: { class_name: className } }),
}

// ============================================================
// 工具
// ============================================================

export const utilApi = {
  openLocation: (data) => http.post('/open-location', data),
}

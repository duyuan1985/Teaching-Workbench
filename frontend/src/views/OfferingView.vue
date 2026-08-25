<template>
  <div class="page-container">
    <!-- 面包屑导航 -->
    <el-breadcrumb separator=">" class="breadcrumb">
      <el-breadcrumb-item :to="{ path: '/' }">返回首页</el-breadcrumb-item>
      <el-breadcrumb-item>课程详情</el-breadcrumb-item>
    </el-breadcrumb>

    <div v-loading="pageLoading">
      <!-- 课程基本信息 -->
      <div class="panel" v-if="offering">
        <h2>
          <el-icon><Document /></el-icon>
          {{ offering.course_name }}
          <el-tag size="small" type="info" class="ml">{{ offering.term }}</el-tag>
          <el-tag v-if="offering.offering_kind === '实训课程'" size="small" type="warning" class="ml">实训</el-tag>
          <el-button size="small" :icon="EditPen" class="ml" @click="openEditDialog">编辑</el-button>
        </h2>
        <el-descriptions :column="3" border>
          <el-descriptions-item label="课程名">{{ offering.course_name }}</el-descriptions-item>
          <el-descriptions-item label="学期">{{ offering.term }}</el-descriptions-item>
          <el-descriptions-item label="专业">{{ offering.major || '—' }}</el-descriptions-item>
          <el-descriptions-item label="班级">{{ offering.teaching_class || '—' }}</el-descriptions-item>
          <el-descriptions-item label="课程编号">{{ offering.course_code || '—' }}</el-descriptions-item>
          <el-descriptions-item label="课程性质">{{ offering.course_nature || '—' }}</el-descriptions-item>
          <el-descriptions-item label="课程类型">{{ offering.course_type || '—' }}</el-descriptions-item>
          <el-descriptions-item label="考核类型">{{ offering.assessment_type || '—' }}</el-descriptions-item>
          <el-descriptions-item label="考核方式">{{ offering.assessment_method || '—' }}</el-descriptions-item>
          <el-descriptions-item label="学分">{{ offering.credits }}</el-descriptions-item>
          <el-descriptions-item label="总学时">{{ offering.total_hours }}</el-descriptions-item>
          <el-descriptions-item label="周学时">{{ offering.weekly_hours }}</el-descriptions-item>
          <el-descriptions-item label="教材版本">{{ offering.textbook_version || '—' }}</el-descriptions-item>
          <el-descriptions-item label="课程类别">{{ offering.offering_kind || '普通课程' }}</el-descriptions-item>
          <el-descriptions-item label="备注">{{ offering.notes || '—' }}</el-descriptions-item>
        </el-descriptions>
      </div>

      <!-- 工作流进度条 -->
      <div class="panel" v-if="workflow.length">
        <h2><el-icon><Operation /></el-icon> 工作流进度</h2>
        <div class="workflow-steps">
          <div
            v-for="step in workflow"
            :key="step.step"
            class="workflow-step"
            :class="{ done: step.done, current: step.step === currentStep && !step.done }"
          >
            <span class="step-num">{{ step.step }}</span>
            {{ step.name }}
          </div>
        </div>
        <div class="counts-row" v-if="counts">
          <el-tag size="small" type="info">排课 {{ counts.sessions }}</el-tag>
          <el-tag size="small" type="info">资源 {{ counts.resources }}</el-tag>
          <el-tag size="small" type="info">蓝本单元 {{ counts.units }}</el-tag>
          <el-tag size="small" type="info">内容模型 {{ counts.models }}</el-tag>
          <el-tag size="small" type="info">草稿 {{ counts.drafts }}</el-tag>
          <el-tag size="small" type="info">文档 {{ counts.documents }}</el-tag>
          <el-tag size="small" type="info">任务 {{ counts.tasks }}</el-tag>
          <el-tag size="small" type="info">模板文件 {{ counts.template_files }}</el-tag>
          <el-tag size="small" type="info">源文件 {{ counts.source_files }}</el-tag>
        </div>
      </div>

      <!-- 变更感知面板 -->
      <div class="panel" v-if="dirtyFlags !== null">
        <h2>
          <el-icon><WarnTriangleFilled /></el-icon>
          变更状态
          <el-badge v-if="dirtyActiveCount > 0" :value="dirtyActiveCount" class="ml" />
        </h2>

        <!-- 有待处理变更 -->
        <el-alert
          v-if="dirtyActiveCount > 0 && dirtyRecommended"
          :title="`检测到 ${dirtyActiveCount} 项变更，建议优先处理：${dirtyRecommended.label}`"
          type="warning"
          :closable="false"
          show-icon
          style="margin-bottom: 12px"
        >
          <template #default>
            <div class="dirty-recommend">
              <span>{{ dirtyRecommended.hint }}</span>
              <el-button
                size="small"
                type="warning"
                round
                style="margin-left: 12px"
                @click="handleRecommendedAction"
              >立即处理</el-button>
            </div>
          </template>
        </el-alert>

        <!-- 变更详情列表 -->
        <el-table :data="dirtyFlags" stripe size="small" style="width: 100%">
          <el-table-column label="数据项" width="160">
            <template #default="{ row }">
              <div style="display: flex; align-items: center; gap: 6px">
                <el-icon v-if="row.active" color="#E6A23C"><WarnTriangleFilled /></el-icon>
                <el-icon v-else color="#67C23A"><CircleCheckFilled /></el-icon>
                <span>{{ row.label }}</span>
              </div>
            </template>
          </el-table-column>
          <el-table-column label="状态" width="90" align="center">
            <template #default="{ row }">
              <el-tag v-if="row.active" size="small" type="warning">待处理</el-tag>
              <el-tag v-else size="small" type="success">已同步</el-tag>
            </template>
          </el-table-column>
          <el-table-column label="变更说明" min-width="180" show-overflow-tooltip>
            <template #default="{ row }">
              <span v-if="row.active" class="dirty-reason">{{ row.reason || '已修改' }}</span>
              <span v-else class="dirty-clean">无需操作</span>
            </template>
          </el-table-column>
          <el-table-column label="影响与建议" min-width="200" show-overflow-tooltip>
            <template #default="{ row }">
              <span v-if="row.active">{{ row.hint }}</span>
              <span v-else class="dirty-clean">该项未变更，无需重复操作</span>
            </template>
          </el-table-column>
          <el-table-column label="操作" width="120" align="center">
            <template #default="{ row }">
              <el-button
                v-if="row.active"
                size="small"
                type="primary"
                link
                @click="handleDirtyAction(row)"
              >去处理</el-button>
              <span v-else class="dirty-clean">—</span>
            </template>
          </el-table-column>
        </el-table>

        <!-- 无变更提示 -->
        <el-alert
          v-if="dirtyActiveCount === 0"
          title="所有数据均为最新，无需重新生成文档"
          type="success"
          :closable="false"
          show-icon
          style="margin-top: 12px"
        />
      </div>

      <!-- 操作按钮区域 -->
      <div class="panel">
        <h2><el-icon><Tools /></el-icon> 工作流操作</h2>

        <!-- 实训课程操作 -->
        <div v-if="isTrainingCourse" class="actions">
          <el-button
            type="primary"
            plain
            :icon="Refresh"
            :loading="loadingMap.rebuildSchedule"
            @click="confirmAction('rebuildSchedule', '重建排课', '确定要重建排课吗？')"
          >重建排课</el-button>
          <el-button
            type="success"
            :icon="Document"
            :loading="loadingMap.generateDocuments"
            @click="confirmAction('generateDocuments', '生成实训资料', '确定要生成实训资料吗？')"
          >生成实训资料</el-button>
          <el-button
            type="danger"
            plain
            :icon="RefreshLeft"
            :loading="loadingMap.resetWorkflow"
            @click="confirmAction('resetWorkflow', '重置流程', '确定要重置工作流吗？将清除蓝本、任务、模型、草稿和已生成文档，此操作不可撤销。')"
          >重置流程</el-button>
        </div>

        <!-- 普通课程操作 -->
        <div v-else class="actions">
          <el-button
            type="primary"
            plain
            :icon="Refresh"
            :loading="loadingMap.rebuildSchedule"
            @click="confirmAction('rebuildSchedule', '重建排课', '确定要重建排课吗？')"
          >重建排课</el-button>
          <el-button
            type="primary"
            plain
            :icon="Refresh"
            :loading="loadingMap.rebuildResources"
            @click="confirmAction('rebuildResources', '重建资源索引', '确定要重建资源索引吗？')"
          >重建资源索引</el-button>
          <el-button
            type="primary"
            plain
            :icon="Refresh"
            :loading="loadingMap.rebuildReview"
            @click="confirmAction('rebuildReview', '重建蓝本审查', '确定要重建蓝本审查吗？')"
          >重建蓝本审查</el-button>
          <el-button
            type="warning"
            plain
            :icon="Select"
            :loading="loadingMap.approveUnits"
            @click="confirmAction('approveUnits', '批量确认蓝本', '确定要批量确认所有蓝本单元吗？')"
          >批量确认蓝本</el-button>
          <el-button
            type="success"
            plain
            :icon="List"
            :loading="loadingMap.buildTasks"
            @click="confirmAction('buildTasks', '构建任务', '确定要构建教学任务吗？')"
          >构建任务</el-button>
          <el-button
            type="primary"
            plain
            :icon="Refresh"
            :loading="loadingMap.rebuildFoundation"
            @click="confirmAction('rebuildFoundation', '重建生成基础', '确定要重建生成基础吗？')"
          >重建生成基础</el-button>
          <el-button
            type="success"
            plain
            :icon="Document"
            :loading="loadingMap.generateDocuments"
            @click="openGenerateDialog"
          >生成文档</el-button>
          <el-button
            type="danger"
            plain
            :icon="RefreshLeft"
            :loading="loadingMap.resetWorkflow"
            @click="confirmAction('resetWorkflow', '重置流程', '确定要重置工作流吗？将清除蓝本、任务、模型、草稿和已生成文档，此操作不可撤销。')"
          >重置流程</el-button>
        </div>

        <!-- 重建基础进度面板 -->
        <div v-if="foundationStages" class="foundation-progress">
          <div class="foundation-progress-header">
            <span v-if="foundationRunning" class="foundation-progress-title">
              <el-icon class="is-loading"><Loading /></el-icon>
              重建生成基础进行中…
            </span>
            <span v-else-if="foundationError" class="foundation-progress-title foundation-error-text">
              <el-icon><CircleCloseFilled /></el-icon>
              重建生成基础失败
            </span>
            <span v-else class="foundation-progress-title foundation-done-text">
              <el-icon><CircleCheckFilled /></el-icon>
              重建生成基础完成
            </span>
          </div>
          <div class="foundation-stages">
            <div
              v-for="stage in Object.keys(FOUNDATION_STAGE_LABELS)"
              :key="stage"
              class="foundation-stage"
              :class="foundationStages[stage]?.status || 'pending'"
            >
              <span class="foundation-stage-icon">{{ FOUNDATION_STAGE_ICONS[stage] }}</span>
              <span class="foundation-stage-name">{{ FOUNDATION_STAGE_LABELS[stage] }}</span>
              <span class="foundation-stage-status">
                <template v-if="foundationStages[stage]?.status === 'running'">进行中</template>
                <template v-else-if="foundationStages[stage]?.status === 'done'">完成</template>
                <template v-else-if="foundationStages[stage]?.status === 'failed'">失败</template>
                <template v-else>等待</template>
              </span>
            </div>
          </div>
          <div v-if="foundationError" class="foundation-error-detail">{{ foundationError }}</div>
        </div>
      </div>

      <!-- 排课列表 -->
      <div class="panel">
        <h2><el-icon><Calendar /></el-icon> 排课列表</h2>
        <el-table :data="sessions" v-loading="sessionsLoading" stripe empty-text="暂无排课数据" style="width: 100%">
          <el-table-column prop="week_no" label="周次" width="70" align="center" />
          <el-table-column prop="lesson_date" label="上课日期" width="120" />
          <el-table-column prop="weekday" label="星期" width="70" align="center" />
          <el-table-column prop="periods" label="节次" width="100" />
          <el-table-column prop="hours" label="学时" width="60" align="center" />
          <el-table-column prop="class_name" label="班级" width="130" show-overflow-tooltip />
          <el-table-column prop="classroom" label="教室" width="120" show-overflow-tooltip />
          <el-table-column prop="session_type" label="类型" width="90" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="sessionTypeTag(row.session_type)">{{ row.session_type }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="status" label="状态" width="80" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="statusTag(row.status)">{{ row.status }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column label="操作" width="70" align="center" fixed="right">
            <template #default="{ row }">
              <el-button size="small" link type="primary" @click="openSessionEdit(row)">编辑</el-button>
            </template>
          </el-table-column>
        </el-table>
      </div>

      <!-- 教材蓝本审查 -->
      <div class="panel" v-if="!isTrainingCourse">
        <h2><el-icon><Reading /></el-icon> 教材蓝本审查</h2>
        <el-table :data="units" v-loading="unitsLoading" stripe empty-text="暂无蓝本单元数据" style="width: 100%">
          <el-table-column prop="seq" label="序号" width="60" align="center" />
          <el-table-column prop="project_title" label="项目名称" min-width="180" show-overflow-tooltip />
          <el-table-column prop="source_file" label="来源文件" min-width="140" show-overflow-tooltip />
          <el-table-column prop="review_action" label="审查动作" width="90" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="reviewActionTag(row.review_action)">{{ row.review_action }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="suggested_hours" label="建议学时" width="90" align="center" />
          <el-table-column prop="revised_focus" label="修订重点" min-width="160" show-overflow-tooltip>
            <template #default="{ row }">
              <span :style="{ color: row.content_warnings && row.content_warnings.length ? '#E8463A' : '' }">{{ row.revised_focus || '—' }}</span>
            </template>
          </el-table-column>
          <el-table-column label="过时警告" width="100" align="center">
            <template #default="{ row }">
              <el-tooltip v-if="row.content_warnings && row.content_warnings.length" placement="bottom" :width="350">
                <template #content>
                  <div v-for="w in row.content_warnings" :key="w.id" style="margin-bottom: 8px">
                    <strong style="color: #FDA4AF">{{ w.topic }}</strong>
                    <el-tag size="small" type="danger" style="margin-left: 4px">{{ w.update_type }}</el-tag>
                    <div style="color: #FECACA; font-size: 12px; margin-top: 2px">{{ w.reason }}</div>
                  </div>
                </template>
                <el-badge :value="row.content_warnings.length" type="danger">
                  <el-icon color="#E8463A" size="18"><Warning /></el-icon>
                </el-badge>
              </el-tooltip>
              <span v-else style="color: #909399">—</span>
            </template>
          </el-table-column>
          <el-table-column prop="approval_status" label="审批状态" width="90" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="approvalTag(row.approval_status)">{{ row.approval_status }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column label="操作" width="140" align="center">
            <template #default="{ row }">
              <el-button v-if="row.approval_status !== '已确认'" size="small" type="success" link @click="approveOneUnit(row)">确认</el-button>
              <el-button v-if="row.approval_status === '已确认'" size="small" type="warning" link @click="rejectOneUnit(row)">退回</el-button>
            </template>
          </el-table-column>
        </el-table>
      </div>

      <!-- 内容模型审查 -->
      <div class="panel" v-if="!isTrainingCourse">
        <h2>
          <el-icon><MagicStick /></el-icon>
          内容模型审查
          <el-tag v-if="contentModelStatus === '已确认'" size="small" type="success" class="ml">已确认</el-tag>
          <el-tag v-else-if="contentModelStatus" size="small" type="warning" class="ml">待检查</el-tag>
          <el-tag v-else size="small" type="info" class="ml">未生成</el-tag>
          <el-button
            v-if="contentModelStatus"
            size="small"
            type="primary"
            class="ml"
            @click="openModelReview"
          >{{ contentModelStatus === '已确认' ? '查看/修改' : '审查并纠正' }}</el-button>
          <span v-else class="model-hint">点击「重建生成基础」后系统将从教材资源自动生成</span>
        </h2>
        <el-alert
          v-if="contentModelStatus && contentModelStatus !== '已确认'"
          title="系统已从教材资源生成课程定位、岗位方向、课程目标和教师要求草案，请审查修改后确认，确认后才能生成正式文档"
          type="info"
          :closable="false"
          show-icon
        />
      </div>

      <!-- 教学任务列表 -->
      <div class="panel" v-if="!isTrainingCourse">
        <h2><el-icon><List /></el-icon> 教学任务列表</h2>
        <el-table :data="tasks" v-loading="tasksLoading" stripe empty-text="暂无教学任务数据" style="width: 100%">
          <el-table-column prop="seq" label="序号" width="60" align="center" />
          <el-table-column prop="chapter" label="章节" width="100" show-overflow-tooltip />
          <el-table-column prop="title" label="任务名称" min-width="180" show-overflow-tooltip />
          <el-table-column prop="hours" label="总学时" width="70" align="center" />
          <el-table-column label="理论/实践" width="90" align="center">
            <template #default="{ row }">
              {{ row.theory_hours }}/{{ row.practice_hours }}
            </template>
          </el-table-column>
          <el-table-column prop="week_no" label="周次" width="60" align="center" />
          <el-table-column prop="lesson_date" label="上课日期" width="110" />
          <el-table-column label="知识目标" min-width="160" show-overflow-tooltip>
            <template #default="{ row }">
              {{ row.knowledge_goal || '—' }}
            </template>
          </el-table-column>
          <el-table-column label="能力目标" min-width="160" show-overflow-tooltip>
            <template #default="{ row }">
              {{ row.ability_goal || '—' }}
            </template>
          </el-table-column>
        </el-table>
      </div>

      <!-- 内容更新建议 -->
      <div class="panel">
        <h2>
          <el-icon><Refresh /></el-icon> 教材内容更新
          <el-button size="small" type="primary" :loading="analyzing" style="margin-left: 12px" @click="analyzeUpdates">
            <el-icon><MagicStick /></el-icon> AI检测过时内容
          </el-button>
          <el-button size="small" style="margin-left: 8px" @click="loadContentUpdates">
            <el-icon><Refresh /></el-icon> 刷新
          </el-button>
        </h2>
        <p class="panel-desc">
          AI 自动分析教材内容，识别过时技术和需要补充的知识点。审核通过后，更新内容会同步融入课程标准、授课计划和教学设计。
        </p>
        <el-table :data="contentUpdates" v-loading="updatesLoading" stripe empty-text="暂无更新建议，点击「AI检测过时内容」开始分析" style="width: 100%">
          <el-table-column prop="topic" label="知识点主题" min-width="150" show-overflow-tooltip />
          <el-table-column prop="update_type" label="类型" width="95" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="updateTypeTag(row.update_type)">{{ row.update_type }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="confidence" label="置信度" width="80" align="center">
            <template #default="{ row }">
              <el-progress :percentage="Math.round(row.confidence * 100)" :stroke-width="6" :show-text="false" />
            </template>
          </el-table-column>
          <el-table-column label="建议更新内容" min-width="240" show-overflow-tooltip>
            <template #default="{ row }">{{ row.suggested_content }}</template>
          </el-table-column>
          <el-table-column label="原因" min-width="150" show-overflow-tooltip>
            <template #default="{ row }">{{ row.reason }}</template>
          </el-table-column>
          <el-table-column prop="status" label="状态" width="80" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="updateStatusTag(row.status)">{{ row.status }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column label="操作" width="170" align="center" fixed="right">
            <template #default="{ row }">
              <el-button
                v-if="row.status === '待审核'"
                size="small"
                type="success"
                @click="reviewUpdate(row, '已采纳')"
              >采纳</el-button>
              <el-button
                v-if="row.status === '待审核'"
                size="small"
                type="info"
                @click="reviewUpdate(row, '已忽略')"
              >忽略</el-button>
              <el-button
                v-if="row.status !== '待审核'"
                size="small"
                @click="reviewUpdate(row, '待审核')"
              >撤回</el-button>
              <el-button size="small" type="danger" text @click="deleteUpdate(row)">删</el-button>
            </template>
          </el-table-column>
        </el-table>
      </div>

      <!-- 已生成文档 -->
      <div class="panel">
        <h2><el-icon><FolderOpened /></el-icon> 已生成文档</h2>
        <el-table :data="documents" v-loading="documentsLoading" stripe empty-text="暂无已生成文档" style="width: 100%">
          <el-table-column prop="document_type" label="文档类型" min-width="160" show-overflow-tooltip />
          <el-table-column prop="output_path" label="输出路径" min-width="240" show-overflow-tooltip />
          <el-table-column prop="generation_status" label="生成状态" width="90" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="docStatusTag(row.generation_status)">{{ row.generation_status }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="structural_check" label="结构检查" width="90" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="checkTag(row.structural_check)" effect="plain">{{ row.structural_check }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="visual_check" label="视觉检查" width="90" align="center">
            <template #default="{ row }">
              <el-tag size="small" :type="checkTag(row.visual_check)" effect="plain">{{ row.visual_check }}</el-tag>
            </template>
          </el-table-column>
          <el-table-column prop="generated_at" label="生成时间" width="160" />
        </el-table>
      </div>
    </div>

    <!-- 生成文档对话框 -->
    <el-dialog v-model="generateDialogVisible" title="生成文档" width="540px" :close-on-click-modal="false">
      <el-form label-width="96px">
        <el-form-item label="生成内容">
          <el-checkbox-group v-model="generateForm.documentTypes">
            <el-checkbox value="课程标准">课程标准</el-checkbox>
            <el-checkbox value="授课计划">授课计划</el-checkbox>
            <el-checkbox value="教学设计">教学设计</el-checkbox>
          </el-checkbox-group>
        </el-form-item>
        <el-form-item v-if="designSelected" label="教学设计范围">
          <el-radio-group v-model="generateForm.designScope">
            <el-radio value="all">全部周次（整本）</el-radio>
            <el-radio value="week">指定周次</el-radio>
          </el-radio-group>
        </el-form-item>
        <el-form-item v-if="designSelected && generateForm.designScope === 'week'" label="选择周次">
          <el-select v-model="generateForm.weekNo" placeholder="请选择周次" style="width: 100%">
            <el-option
              v-for="w in weekOptions"
              :key="w.week"
              :label="`第${w.week}周（${w.taskCount}个任务）`"
              :value="w.week"
            />
          </el-select>
          <div v-if="weekTaskSummary" class="week-hint">{{ weekTaskSummary }}</div>
        </el-form-item>
      </el-form>

      <!-- 生成条件（按所选文档范围检查） -->
      <div v-if="!isTrainingCourse" class="readiness-box" :class="readinessBlockers.length ? 'readiness-warn' : 'readiness-ok'">
        <template v-if="readinessLoading">正在检查生成条件...</template>
        <template v-else-if="!readinessBlockers.length">生成条件已满足，可以开始生成。</template>
        <template v-else>
          <div class="readiness-title">生成条件不满足（{{ readinessBlockers.length }}项）：</div>
          <div v-for="blocker in readinessBlockers" :key="blocker.key" class="readiness-item">
            <span class="readiness-text">{{ blocker.text }}</span>
            <el-button
              v-if="blocker.confirmType"
              size="small"
              type="primary"
              :loading="confirmingKey === blocker.key"
              @click="confirmReadiness(blocker)"
            >{{ blocker.confirmLabel }}</el-button>
          </div>
        </template>
      </div>
      <template #footer>
        <el-button @click="generateDialogVisible = false">取消</el-button>
        <el-button
          type="success"
          :loading="loadingMap.generateDocuments"
          :disabled="!isTrainingCourse && readinessBlockers.length > 0"
          @click="submitGenerate"
        >开始生成</el-button>
      </template>
    </el-dialog>
    <!-- 编辑课程基本信息对话框 -->
    <el-dialog v-model="editDialogVisible" title="编辑课程基本信息" width="720px" :close-on-click-modal="false">
      <el-form :model="editForm" label-width="90px">
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="课程名" required>
              <el-input v-model="editForm.course_name" placeholder="课程名称" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="课程编号">
              <el-input v-model="editForm.course_code" placeholder="课程编号" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="专业">
              <el-input v-model="editForm.major" placeholder="专业名称" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="班级">
              <el-input v-model="editForm.teaching_class" placeholder="班级名称" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="课程性质">
              <el-select v-model="editForm.course_nature" placeholder="选择课程性质" clearable style="width: 100%">
                <el-option label="必修课" value="必修课" />
                <el-option label="选修课" value="选修课" />
                <el-option label="限选课" value="限选课" />
              </el-select>
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="课程类型">
              <el-select v-model="editForm.course_type" placeholder="选择课程类型" clearable style="width: 100%">
                <el-option
                  v-for="t in courseTypes"
                  :key="t.id"
                  :label="t.name"
                  :value="t.name"
                />
              </el-select>
              <div class="edit-hint">选项来自「基础设置 → 课程类型管理」</div>
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="8">
            <el-form-item label="学分">
              <el-input-number v-model="editForm.credits" :min="0" :precision="1" :step="0.5" controls-position="right" style="width: 100%" />
            </el-form-item>
          </el-col>
          <el-col :span="8">
            <el-form-item label="总学时">
              <el-input-number v-model="editForm.total_hours" :min="0" :step="2" controls-position="right" style="width: 100%" />
            </el-form-item>
          </el-col>
          <el-col :span="8">
            <el-form-item label="周学时">
              <el-input-number v-model="editForm.weekly_hours" :min="0" :step="1" controls-position="right" style="width: 100%" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="考核类型">
              <el-input v-model="editForm.assessment_type" placeholder="如 期末考核" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="考核方式">
              <el-input v-model="editForm.assessment_method" placeholder="如 实操" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item label="教材版本">
          <el-input v-model="editForm.textbook_version" placeholder="教材版本信息" />
        </el-form-item>
        <el-form-item label="备注">
          <el-input v-model="editForm.notes" type="textarea" :rows="2" placeholder="备注信息" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="editDialogVisible = false">取消</el-button>
        <el-button type="primary" :loading="editSaving" @click="handleEditSave">保存</el-button>
      </template>
    </el-dialog>

    <!-- 编辑排课记录对话框 -->
    <el-dialog v-model="sessionDialogVisible" title="编辑排课记录" width="560px" :close-on-click-modal="false">
      <el-form :model="sessionForm" label-width="80px">
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="类型">
              <el-select v-model="sessionForm.session_type" style="width: 100%">
                <el-option v-for="t in sessionTypeOptions" :key="t" :label="t" :value="t" />
              </el-select>
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="状态">
              <el-select v-model="sessionForm.status" style="width: 100%">
                <el-option v-for="s in sessionStatusOptions" :key="s" :label="s" :value="s" />
              </el-select>
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="周次">
              <el-input-number v-model="sessionForm.week_no" :min="0" :max="30" controls-position="right" style="width: 100%" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="学时">
              <el-input-number v-model="sessionForm.hours" :min="0" :max="12" controls-position="right" style="width: 100%" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item label="上课日期">
          <el-date-picker
            v-model="sessionForm.lesson_date"
            type="date"
            value-format="YYYY-MM-DD"
            placeholder="选择日期（可留空）"
            style="width: 100%"
          />
          <div v-if="sessionWeekdayPreview" class="edit-hint">对应星期：{{ sessionWeekdayPreview }}</div>
        </el-form-item>
        <el-form-item label="节次">
          <el-input v-model="sessionForm.periods" placeholder="如 1-2节" />
        </el-form-item>
        <el-form-item label="教室">
          <el-input v-model="sessionForm.classroom" placeholder="如 802教室" />
        </el-form-item>
        <div v-if="sessionForm.source_note" class="session-note">来源：{{ sessionForm.source_note }}</div>
      </el-form>
      <template #footer>
        <el-button @click="sessionDialogVisible = false">取消</el-button>
        <el-button type="primary" :loading="sessionSaving" @click="handleSessionSave">保存</el-button>
      </template>
    </el-dialog>

    <!-- 内容模型审查对话框 -->
    <el-dialog v-model="modelDialogVisible" title="内容模型审查" width="760px" :close-on-click-modal="false">
      <div v-loading="modelLoading">
        <el-alert
          v-for="(item, i) in modelReview.missing_evidence"
          :key="i"
          :title="item"
          type="warning"
          :closable="false"
          show-icon
          style="margin-bottom: 8px"
        />

        <h3 class="model-section-title">课程定位（来自课程基本信息）</h3>
        <el-descriptions :column="2" border size="small">
          <el-descriptions-item label="课程名称">{{ modelReview.identity.course_name }}</el-descriptions-item>
          <el-descriptions-item label="课程性质">{{ modelReview.identity.course_nature || '—' }}</el-descriptions-item>
          <el-descriptions-item label="课程类型">{{ modelReview.identity.course_type || '—' }}</el-descriptions-item>
          <el-descriptions-item label="学时/学分">{{ modelReview.identity.total_hours }}学时 / {{ modelReview.identity.credits }}学分</el-descriptions-item>
          <el-descriptions-item label="适用专业">{{ modelReview.identity.major || '—' }}</el-descriptions-item>
          <el-descriptions-item label="教材版本">{{ modelReview.identity.textbook_version || '—' }}</el-descriptions-item>
          <el-descriptions-item label="先导课程">{{ modelReview.identity.prerequisite_courses || '—' }}</el-descriptions-item>
          <el-descriptions-item label="后续课程">{{ modelReview.identity.followup_courses || '—' }}</el-descriptions-item>
        </el-descriptions>
        <div class="edit-hint">如需修改，请使用课程基本信息面板的「编辑」按钮</div>

        <template v-if="modelReview.talent_plan">
          <h3 class="model-section-title">人才培养方案依据</h3>
          <el-descriptions :column="3" border size="small">
            <el-descriptions-item label="方案版本">{{ modelReview.talent_plan.cohort || '未标注' }}</el-descriptions-item>
            <el-descriptions-item label="方案中课程代码">{{ modelReview.talent_plan.course_info?.code || '未收录' }}</el-descriptions-item>
            <el-descriptions-item label="方案中课程类别">{{ modelReview.talent_plan.course_info?.category || '未收录' }}</el-descriptions-item>
          </el-descriptions>
          <div v-if="jobPositionText" class="edit-hint">方案职业面向岗位群：{{ jobPositionText }}</div>
          <div class="edit-hint">课程性质、课程设计等内容将引用该方案的培养目标与培养规格生成；方案中未收录本课程时按现有资料推导</div>
        </template>

        <h3 class="model-section-title">岗位方向（可编辑，每行一条）</h3>
        <el-input
          v-model="modelForm.ability_outcomes_text"
          type="textarea"
          :rows="5"
          placeholder="岗位方向与能力成果，每行一条"
        />

        <h3 class="model-section-title">知识体系（来自教材蓝本，供参考）</h3>
        <div class="model-tags">
          <el-tag v-for="(item, i) in modelReview.knowledge_system" :key="i" size="small" class="model-tag">{{ item }}</el-tag>
          <span v-if="!modelReview.knowledge_system.length" class="dirty-clean">暂无</span>
        </div>

        <h3 class="model-section-title">技术工具（来自教材蓝本，供参考）</h3>
        <div class="model-tags">
          <el-tag v-for="(item, i) in modelReview.tools_technology" :key="i" size="small" class="model-tag" type="info">{{ item }}</el-tag>
          <span v-if="!modelReview.tools_technology.length" class="dirty-clean">暂无</span>
        </div>

        <h3 class="model-section-title">教学方法（可编辑，每行一条）</h3>
        <el-input
          v-model="modelForm.teaching_methods_text"
          type="textarea"
          :rows="3"
          placeholder="教学方法，每行一条"
        />
      </div>
      <template #footer>
        <el-button @click="modelDialogVisible = false">取消</el-button>
        <el-button :loading="modelSaving" @click="handleModelSave(false)">仅保存修改</el-button>
        <el-button type="success" :loading="modelSaving" @click="handleModelSave(true)">保存并确认通过</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, computed, onMounted, watch } from 'vue'
import { useRoute } from 'vue-router'
import { ElMessage, ElMessageBox, ElNotification } from 'element-plus'
import {
  Document, Operation, Tools, Refresh, RefreshLeft,
  Calendar, Reading, List, FolderOpened, Select,
  WarnTriangleFilled, CircleCheckFilled, CircleCloseFilled, MagicStick,
  Warning, EditPen, Loading,
} from '@element-plus/icons-vue'
import { offeringsApi, courseTypesApi, sessionsApi } from '../api'

const route = useRoute()
const offeringId = computed(() => route.params.id)

// ---- 页面数据 ----
const pageLoading = ref(false)
const offering = ref(null)
const workflow = ref([])
const currentStep = ref(1)
const counts = ref(null)

// ---- 变更感知 ----
const dirtyFlags = ref(null)
const dirtyActiveCount = computed(() => {
  if (!dirtyFlags.value) return 0
  return dirtyFlags.value.filter(f => f.active).length
})
const dirtyRecommended = computed(() => {
  if (!dirtyFlags.value) return null
  return dirtyFlags.value.find(f => f.active) || null
})

const isTrainingCourse = computed(() => {
  return offering.value && offering.value.offering_kind === '实训课程'
})

async function loadDirtyFlags() {
  try {
    const data = await offeringsApi.dirtyFlags(offeringId)
    dirtyFlags.value = data.flags
  } catch {
    // 错误已由 axios 拦截器提示
  }
}

const dirtyActionMap = {
  generate: { key: 'generateDocuments', label: '生成文档' },
  rebuild_review: { key: 'rebuildReview', label: '重建蓝本审查' },
  build_tasks: { key: 'buildTasks', label: '构建任务' },
  rebuild_foundation: { key: 'rebuildFoundation', label: '重建生成基础' },
}

async function handleDirtyAction(row) {
  const action = dirtyActionMap[row.action]
  if (!action) return
  if (action.key === 'generateDocuments' && !isTrainingCourse.value) {
    openGenerateDialog()
    return
  }
  await confirmAction(action.key, action.label, `${row.label}已变更，确定要执行「${action.label}」吗？`)
}

async function handleRecommendedAction() {
  if (!dirtyRecommended.value) return
  await handleDirtyAction(dirtyRecommended.value)
}

async function loadOffering() {
  pageLoading.value = true
  try {
    const data = await offeringsApi.get(offeringId)
    offering.value = data.offering
    workflow.value = data.workflow || []
    currentStep.value = data.current_step || 1
    counts.value = data.counts || null
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    pageLoading.value = false
  }
}

// ---- 排课列表 ----
const sessions = ref([])
const sessionsLoading = ref(false)

async function loadSessions() {
  sessionsLoading.value = true
  try {
    sessions.value = await offeringsApi.sessions(offeringId)
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    sessionsLoading.value = false
  }
}

// ---- 蓝本单元 ----
const units = ref([])
const unitsLoading = ref(false)

async function loadUnits() {
  unitsLoading.value = true
  try {
    units.value = await offeringsApi.units(offeringId)
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    unitsLoading.value = false
  }
}

async function approveOneUnit(row) {
  if (row.content_warnings && row.content_warnings.length) {
    const topics = row.content_warnings.map(w => w.topic).join('、')
    try {
      await ElMessageBox.confirm(
        `「${row.project_title}」的修订重点包含 ${row.content_warnings.length} 条过时内容警告：\n${topics}\n\n建议先到「教材内容更新」面板处理这些警告后再确认。\n\n是否仍然强制确认？`,
        '过时内容警告',
        { type: 'warning', confirmButtonText: '仍然确认', cancelButtonText: '去处理' }
      )
    } catch {
      return
    }
  }
  try {
    await offeringsApi.updateUnit(row.id, { approval_status: '已确认' })
    ElMessage.success(`「${row.project_title}」已确认`)
    await loadUnits()
    await loadDirtyFlags()
  } catch {
    // 错误已由拦截器提示
  }
}

async function rejectOneUnit(row) {
  try {
    await offeringsApi.updateUnit(row.id, { approval_status: '待确认' })
    ElMessage.success(`「${row.project_title}」已退回修改`)
    await loadUnits()
    await loadDirtyFlags()
  } catch {
    // 错误已由拦截器提示
  }
}

// ---- 教学任务 ----
const tasks = ref([])
const tasksLoading = ref(false)

async function loadTasks() {
  tasksLoading.value = true
  try {
    tasks.value = await offeringsApi.tasks(offeringId)
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    tasksLoading.value = false
  }
}

// ---- 已生成文档 ----
const documents = ref([])
const documentsLoading = ref(false)

async function loadDocuments() {
  documentsLoading.value = true
  try {
    documents.value = await offeringsApi.documents(offeringId)
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    documentsLoading.value = false
  }
}

// ---- 操作按钮 ----
const loadingMap = reactive({
  rebuildSchedule: false,
  rebuildResources: false,
  rebuildReview: false,
  rebuildFoundation: false,
  generateDocuments: false,
  buildTasks: false,
  approveUnits: false,
  resetWorkflow: false,
})

// 重建基础进度（页面内联显示，不依赖通知弹窗）
const foundationStages = ref(null)
const foundationRunning = ref(false)
const foundationError = ref('')

const FOUNDATION_STAGE_ICONS = {
  resources: '📋',
  templates: '📐',
  tasks: '✏️',
  model: '🧠',
  content: '✍️',
  quality: '✅',
}

async function loadFoundationStatus() {
  try {
    const status = await offeringsApi.foundationStatus(offeringId)
    foundationStages.value = status.stages || {}
    foundationRunning.value = !!status.running
    foundationError.value = status.error || ''
    return status
  } catch {
    return null
  }
}

const actionConfig = {
  rebuildSchedule: { label: '重建排课', api: () => offeringsApi.rebuildSchedule(offeringId), reload: ['sessions', 'offering', 'dirtyFlags'] },
  rebuildResources: { label: '重建资源索引', api: () => offeringsApi.rebuildResources(offeringId), reload: ['offering', 'dirtyFlags'] },
  rebuildReview: { label: '重建蓝本审查', api: () => offeringsApi.rebuildReview(offeringId), reload: ['units', 'offering', 'dirtyFlags'] },
  rebuildFoundation: { label: '重建生成基础', api: () => offeringsApi.rebuildFoundation(offeringId), reload: ['offering', 'dirtyFlags'] },
  generateDocuments: { label: '生成文档', api: () => offeringsApi.generateDocuments(offeringId), reload: ['documents', 'offering', 'dirtyFlags'] },
  buildTasks: { label: '构建任务', api: () => offeringsApi.buildTasks(offeringId), reload: ['tasks', 'offering', 'dirtyFlags'] },
  approveUnits: { label: '批量确认蓝本', api: () => offeringsApi.approveUnits(offeringId), reload: ['units', 'offering', 'dirtyFlags'] },
  resetWorkflow: { label: '重置流程', api: () => offeringsApi.resetWorkflow(offeringId), reload: ['units', 'tasks', 'documents', 'offering', 'dirtyFlags'] },
}

const reloadMap = {
  offering: loadOffering,
  sessions: loadSessions,
  units: loadUnits,
  tasks: loadTasks,
  documents: loadDocuments,
  dirtyFlags: loadDirtyFlags,
}

// 重建基础含AI内容生成，后端为后台任务，前端轮询进度直到完成
const FOUNDATION_STAGE_LABELS = {
  resources: '资源解析',
  templates: '模板分析',
  tasks: '任务增强',
  model: '语义模型',
  content: '内容生成',
  quality: '质量检查',
}

function foundationStageText(stages) {
  const order = Object.keys(FOUNDATION_STAGE_LABELS)
  const runningStage = order.find(s => stages[s]?.status === 'running')
  if (runningStage) return `正在${FOUNDATION_STAGE_LABELS[runningStage]}…`
  const failedStage = order.find(s => stages[s]?.status === 'failed')
  if (failedStage) return `${FOUNDATION_STAGE_LABELS[failedStage]}失败`
  const doneCount = order.filter(s => stages[s]?.status === 'done').length
  return doneCount ? `已完成 ${doneCount}/${order.length} 阶段` : '启动中…'
}

async function pollFoundationProgress() {
  const POLL_MS = 3000
  const MAX_WAIT_MS = 30 * 60 * 1000
  const startAt = Date.now()
  let notify = null
  let lastText = ''
  let pollFails = 0
  while (Date.now() - startAt < MAX_WAIT_MS) {
    await new Promise(resolve => setTimeout(resolve, POLL_MS))
    let status
    try {
      status = await offeringsApi.foundationStatus(offeringId)
      pollFails = 0
    } catch {
      pollFails += 1
      if (pollFails >= 5) throw new Error('连续查询进度失败，请刷新页面查看重建结果。')
      continue
    }
    foundationStages.value = status.stages || {}
    foundationRunning.value = !!status.running
    foundationError.value = status.error || ''
    const text = foundationStageText(status.stages || {})
    if (text !== lastText) {
      lastText = text
      if (notify) notify.close()
      notify = ElNotification({ title: '重建生成基础', message: text, duration: 0, type: 'info' })
    }
    if (!status.running) {
      if (notify) notify.close()
      return status
    }
  }
  if (notify) notify.close()
  throw new Error('重建耗时过长，后台仍在执行，请稍后刷新页面查看结果。')
}

async function confirmAction(key, title, message) {
  if (key === 'approveUnits') {
    const warningUnits = units.value.filter(u => u.content_warnings && u.content_warnings.length)
    if (warningUnits.length > 0) {
      const topics = warningUnits.flatMap(u => u.content_warnings.map(w => w.topic)).join('、')
      message += `\n\n注意：${warningUnits.length} 个单元包含过时内容警告（${topics}），建议先到「教材内容更新」面板处理。`
    }
  }
  try {
    await ElMessageBox.confirm(message, title, {
      confirmButtonText: '确定',
      cancelButtonText: '取消',
      type: 'warning',
    })
  } catch {
    return
  }
  loadingMap[key] = true
  try {
    const config = actionConfig[key]
    if (key === 'rebuildFoundation') {
      await offeringsApi.rebuildFoundation(offeringId)
      const status = await pollFoundationProgress()
      if (status.error) {
        ElMessage.error(`重建生成基础失败：${status.error}`)
        return
      }
    } else {
      await config.api()
    }
    ElMessage.success(`${config.label}完成`)
    for (const target of config.reload) {
      await reloadMap[target]()
    }
  } catch (e) {
    if (e instanceof Error && e.message && !e.response && !e.config) {
      ElMessage.error(e.message)
    }
    // 其余错误已由 axios 拦截器提示
  } finally {
    loadingMap[key] = false
  }
}

// ---- 生成文档（可选文件与周次） ----
const generateDialogVisible = ref(false)
const generateForm = reactive({
  documentTypes: ['课程标准', '授课计划', '教学设计'],
  designScope: 'all',
  weekNo: null,
})

const designSelected = computed(() => generateForm.documentTypes.includes('教学设计'))

const weekOptions = computed(() => {
  const counts = new Map()
  for (const t of tasks.value) {
    if (t.week_no == null) continue
    counts.set(t.week_no, (counts.get(t.week_no) || 0) + 1)
  }
  return [...counts.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([week, taskCount]) => ({ week, taskCount }))
})

const weekTaskSummary = computed(() => {
  if (generateForm.weekNo == null) return ''
  const weekTasks = tasks.value.filter(t => t.week_no === generateForm.weekNo)
  if (!weekTasks.length) return ''
  const detail = weekTasks.map(t => `任务${t.seq}（${t.hours}学时）`).join('、')
  return `本周共 ${weekTasks.length} 个任务：${detail}`
})

function openGenerateDialog() {
  generateForm.documentTypes = ['课程标准', '授课计划', '教学设计']
  generateForm.designScope = 'all'
  generateForm.weekNo = null
  generateDialogVisible.value = true
  if (!isTrainingCourse.value) loadReadiness()
}

// ---- 生成条件检查与确认 ----
const readiness = ref(null)
const readinessLoading = ref(false)
const confirmingKey = ref('')

const readinessBlockers = computed(() => {
  if (!readiness.value) return []
  return (readiness.value.blockers || []).map((text) => {
    if (text.includes('尚未审查确认') && readiness.value.content_model) {
      return { key: 'content_model', text, confirmType: 'model', confirmLabel: '去审查并纠正' }
    }
    const match = text.match(/^(课程标准|授课计划|教学设计|实训资料)模板规则尚未确认$/)
    if (match) {
      const tpl = (readiness.value.templates || {})[match[1]]
      if (tpl && tpl.analysis_status !== '未分析') {
        return { key: `template_${match[1]}`, text, confirmType: 'template', templateFileId: tpl.template_file_id, confirmLabel: '确认规则' }
      }
    }
    return { key: text, text }
  })
})

watch(
  () => [...generateForm.documentTypes],
  () => {
    if (generateDialogVisible.value && !isTrainingCourse.value) loadReadiness()
  }
)

async function loadReadiness() {
  readinessLoading.value = true
  try {
    readiness.value = await offeringsApi.generationReadiness(offeringId, generateForm.documentTypes)
  } catch {
    readiness.value = null
  } finally {
    readinessLoading.value = false
  }
}

async function confirmReadiness(blocker) {
  if (blocker.confirmType === 'model') {
    openModelReview()
    return
  }
  confirmingKey.value = blocker.key
  try {
    if (blocker.confirmType === 'template') {
      await offeringsApi.confirmTemplateAnalysis(blocker.templateFileId)
    }
    ElMessage.success('已确认')
    await loadReadiness()
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    confirmingKey.value = ''
  }
}

async function submitGenerate() {
  if (!generateForm.documentTypes.length) {
    ElMessage.warning('请至少选择一个要生成的文档')
    return
  }
  const useWeek = designSelected.value && generateForm.designScope === 'week'
  if (useWeek && generateForm.weekNo == null) {
    ElMessage.warning('请选择要生成的周次')
    return
  }
  const isFull = generateForm.documentTypes.length === 3 && !useWeek
  const scopeText = useWeek
    ? `第${generateForm.weekNo}周单元教学设计`
    : generateForm.documentTypes.join('、')
  try {
    await ElMessageBox.confirm(
      `确定要生成「${scopeText}」吗？${useWeek ? '将生成仅包含该周单元设计的独立文档，不影响整本教学设计。' : ''}`,
      '生成文档',
      { confirmButtonText: '开始生成', cancelButtonText: '取消', type: 'info' }
    )
  } catch {
    return
  }
  loadingMap.generateDocuments = true
  try {
    await offeringsApi.generateDocuments(offeringId, {
      document_types: generateForm.documentTypes,
      week_no: useWeek ? generateForm.weekNo : null,
    })
    ElMessage.success(`${isFull ? '文档' : scopeText}生成完成`)
    generateDialogVisible.value = false
    for (const target of ['documents', 'offering', 'dirtyFlags']) {
      await reloadMap[target]()
    }
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    loadingMap.generateDocuments = false
  }
}

// ---- 课程基本信息编辑 ----
const editDialogVisible = ref(false)
const editSaving = ref(false)
const courseTypes = ref([])
const editForm = reactive({
  course_name: '',
  course_code: '',
  major: '',
  teaching_class: '',
  course_nature: '',
  course_type: '',
  assessment_type: '',
  assessment_method: '',
  credits: 0,
  total_hours: 0,
  weekly_hours: 0,
  textbook_version: '',
  notes: '',
})

async function loadCourseTypes() {
  try {
    courseTypes.value = await courseTypesApi.list()
  } catch {
    // 错误已由 axios 拦截器提示
  }
}

function openEditDialog() {
  const o = offering.value
  if (!o) return
  Object.assign(editForm, {
    course_name: o.course_name || '',
    course_code: o.course_code || '',
    major: o.major || '',
    teaching_class: o.teaching_class || '',
    course_nature: o.course_nature || '',
    course_type: o.course_type || '',
    assessment_type: o.assessment_type || '',
    assessment_method: o.assessment_method || '',
    credits: o.credits ?? 0,
    total_hours: o.total_hours ?? 0,
    weekly_hours: o.weekly_hours ?? 0,
    textbook_version: o.textbook_version || '',
    notes: o.notes || '',
  })
  editDialogVisible.value = true
}

async function handleEditSave() {
  if (!editForm.course_name.trim()) {
    ElMessage.warning('课程名不能为空')
    return
  }
  editSaving.value = true
  try {
    await offeringsApi.update(offeringId, { ...editForm })
    ElMessage.success('课程基本信息已保存，相关文档需重新生成')
    editDialogVisible.value = false
    await loadOffering()
    await loadDirtyFlags()
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    editSaving.value = false
  }
}

// ---- 排课编辑 ----
const sessionDialogVisible = ref(false)
const sessionSaving = ref(false)
const sessionTypeOptions = ['正常排课', '补课', '调课', '停课']
const sessionStatusOptions = ['待确认', '已确认', '已取消']
const sessionForm = reactive({
  id: null,
  week_no: 0,
  lesson_date: '',
  classroom: '',
  status: '待确认',
  session_type: '正常排课',
  periods: '',
  hours: 0,
  source_note: '',
})

const sessionWeekdayPreview = computed(() => {
  if (!sessionForm.lesson_date) return ''
  const day = new Date(`${sessionForm.lesson_date}T00:00:00`).getDay()
  return ['周日', '周一', '周二', '周三', '周四', '周五', '周六'][day]
})

function openSessionEdit(row) {
  Object.assign(sessionForm, {
    id: row.id,
    week_no: row.week_no ?? 0,
    lesson_date: row.lesson_date || '',
    classroom: row.classroom || '',
    status: row.status || '待确认',
    session_type: row.session_type || '正常排课',
    periods: row.periods || '',
    hours: row.hours ?? 0,
    source_note: row.source_note || '',
  })
  sessionDialogVisible.value = true
}

async function handleSessionSave() {
  sessionSaving.value = true
  try {
    await sessionsApi.update(sessionForm.id, {
      week_no: sessionForm.week_no,
      lesson_date: sessionForm.lesson_date,
      classroom: sessionForm.classroom,
      status: sessionForm.status,
      session_type: sessionForm.session_type,
      periods: sessionForm.periods,
      hours: sessionForm.hours,
    })
    ElMessage.success('排课记录已保存')
    sessionDialogVisible.value = false
    await loadSessions()
    await loadDirtyFlags()
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    sessionSaving.value = false
  }
}

// ---- 内容模型审查 ----
const contentModelStatus = ref('')
const modelDialogVisible = ref(false)
const modelLoading = ref(false)
const modelSaving = ref(false)
const modelReview = reactive({
  identity: {},
  knowledge_system: [],
  tools_technology: [],
  missing_evidence: [],
  talent_plan: null,
})
const modelForm = reactive({
  ability_outcomes_text: '',
  teaching_methods_text: '',
})

async function loadContentModelStatus() {
  try {
    const model = await offeringsApi.contentModel(offeringId)
    contentModelStatus.value = model.review_status || '待检查'
  } catch {
    contentModelStatus.value = ''
  }
}

async function openModelReview() {
  modelDialogVisible.value = true
  modelLoading.value = true
  try {
    const model = await offeringsApi.contentModel(offeringId)
    const data = model.model_json || {}
    contentModelStatus.value = model.review_status || '待检查'
    Object.assign(modelReview, {
      identity: data.identity || {},
      knowledge_system: data.knowledge_system || [],
      tools_technology: data.tools_technology || [],
      missing_evidence: data.missing_evidence || [],
      talent_plan: data.talent_plan || null,
    })
    modelForm.ability_outcomes_text = (data.ability_outcomes || []).join('\n')
    modelForm.teaching_methods_text = (data.teaching_methods || []).join('\n')
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    modelLoading.value = false
  }
}

async function handleModelSave(confirm) {
  modelSaving.value = true
  try {
    const ability = modelForm.ability_outcomes_text.split('\n').map(s => s.trim()).filter(Boolean)
    const methods = modelForm.teaching_methods_text.split('\n').map(s => s.trim()).filter(Boolean)
    await offeringsApi.updateContentModel(offeringId, {
      ability_outcomes: ability,
      teaching_methods: methods,
    })
    if (confirm) {
      await offeringsApi.confirmContentModel(offeringId)
      ElMessage.success('内容模型已确认，可以生成正式文档')
    } else {
      ElMessage.success('修改已保存，模型状态重置为待检查')
    }
    modelDialogVisible.value = false
    await loadContentModelStatus()
    if (generateDialogVisible.value) await loadReadiness()
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    modelSaving.value = false
  }
}

const jobPositionText = computed(() => {
  const positions = modelReview.talent_plan?.orientation?.job_positions || []
  return positions.join('、')
})

// ---- 标签颜色 ----
function sessionTypeTag(type) {
  const map = { 正常排课: '', 补课: 'warning', 调课: 'warning', 停课: 'danger' }
  return map[type] || ''
}

function statusTag(status) {
  const map = { 已确认: 'success', 待确认: 'warning', 已取消: 'info' }
  return map[status] || ''
}

function reviewActionTag(action) {
  const map = { 保留: 'info', 更新: '', 补充: 'success', 删除: 'danger' }
  return map[action] || ''
}

function approvalTag(status) {
  const map = { 已确认: 'success', 待确认: 'warning', 退回修改: 'danger' }
  return map[status] || ''
}

function docStatusTag(status) {
  const map = { 已生成: 'success', 草稿: 'warning', 失败: 'danger' }
  return map[status] || ''
}

function checkTag(status) {
  const map = { 通过: 'success', 待检查: 'warning', 不通过: 'danger', 已检查: '' }
  return map[status] || ''
}

// ---- 内容更新建议 ----
const contentUpdates = ref([])
const updatesLoading = ref(false)
const analyzing = ref(false)

function updateTypeTag(type) {
  const map = {
    '技术更新': 'primary',
    '内容补充': 'success',
    '废弃警告': 'danger',
    '最佳实践更新': 'warning',
    '行业趋势': 'info',
  }
  return map[type] || ''
}

function updateStatusTag(status) {
  const map = {
    '待审核': 'warning',
    '已采纳': 'success',
    '已忽略': 'info',
  }
  return map[status] || ''
}

async function loadContentUpdates() {
  updatesLoading.value = true
  try {
    const data = await offeringsApi.contentUpdates(offeringId)
    contentUpdates.value = data.items || []
  } catch (e) {
    ElMessage.error('加载内容更新失败')
  } finally {
    updatesLoading.value = false
  }
}

async function analyzeUpdates() {
  analyzing.value = true
  try {
    const result = await offeringsApi.analyzeContentUpdates(offeringId)
    if (result.new_suggestions > 0) {
      ElMessage.success(`分析完成，发现 ${result.new_suggestions} 条新建议`)
    } else {
      ElMessage.info('分析完成，未发现新的内容更新建议')
    }
    loadContentUpdates()
    loadDirtyFlags()
  } catch (e) {
    ElMessage.error('分析失败：' + (e.response?.data?.detail || e.message))
  } finally {
    analyzing.value = false
  }
}

async function reviewUpdate(row, status) {
  try {
    await offeringsApi.reviewContentUpdate(row.id, status)
    ElMessage.success(status === '已采纳' ? '已采纳，重新生成文档时会自动融入' : status === '已忽略' ? '已忽略' : '已撤回')
    loadContentUpdates()
    loadDirtyFlags()
  } catch (e) {
    ElMessage.error('操作失败')
  }
}

async function deleteUpdate(row) {
  try {
    await ElMessageBox.confirm(`确定删除「${row.topic}」这条更新建议吗？`, '确认删除', { type: 'warning' })
    await offeringsApi.deleteContentUpdate(row.id)
    ElMessage.success('已删除')
    loadContentUpdates()
  } catch (e) {
    if (e !== 'cancel') {
      ElMessage.error('删除失败')
    }
  }
}

onMounted(async () => {
  loadOffering()
  loadSessions()
  loadUnits()
  loadTasks()
  loadDocuments()
  loadDirtyFlags()
  loadContentUpdates()
  loadCourseTypes()
  loadContentModelStatus()
  // 检测后台是否正在执行重建基础，是则恢复轮询
  const status = await loadFoundationStatus()
  if (status && status.running) {
    loadingMap.rebuildFoundation = true
    pollFoundationProgress().then(s => {
      if (s.error) {
        ElMessage.error(`重建生成基础失败：${s.error}`)
      } else {
        ElMessage.success('重建生成基础完成')
      }
    }).catch(e => {
      if (e instanceof Error && e.message) ElMessage.error(e.message)
    }).finally(() => {
      loadingMap.rebuildFoundation = false
    })
  }
})
</script>

<style scoped>
.breadcrumb {
  margin-bottom: 16px;
}

.ml {
  margin-left: 8px;
}

.week-hint {
  margin-top: 6px;
  font-size: 12px;
  color: var(--el-text-color-secondary);
  line-height: 1.5;
}

.readiness-box {
  margin: 4px 0 12px;
  padding: 10px 14px;
  border-radius: 6px;
  font-size: 13px;
  line-height: 1.6;
}

.readiness-ok {
  background: var(--el-color-success-light-9);
  color: var(--el-color-success);
}

.readiness-warn {
  background: var(--el-color-warning-light-9);
  color: var(--el-color-warning-dark-2);
}

.readiness-title {
  font-weight: 600;
  margin-bottom: 4px;
}

.readiness-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 2px 0;
}

.readiness-text {
  flex: 1;
}

.counts-row {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-top: 16px;
}

.panel h2 {
  display: flex;
  align-items: center;
  gap: 8px;
}

.dirty-reason {
  color: #E6A23C;
}

.dirty-clean {
  color: #909399;
}

.dirty-recommend {
  display: flex;
  align-items: center;
  margin-top: 4px;
}

.edit-hint {
  font-size: 12px;
  color: var(--el-text-color-secondary);
  line-height: 1.5;
  margin-top: 4px;
}

.session-note {
  font-size: 12px;
  color: var(--el-text-color-secondary);
  background: var(--el-fill-color-light);
  border-radius: 4px;
  padding: 8px 10px;
  line-height: 1.6;
  margin-top: 8px;
  word-break: break-all;
}

.model-hint {
  font-size: 12px;
  color: var(--el-text-color-secondary);
  margin-left: 8px;
}

.model-section-title {
  font-size: 14px;
  font-weight: 600;
  margin: 18px 0 8px;
}

.model-tags {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}

.model-tag {
  max-width: 100%;
}

/* 重建基础进度面板 */
.foundation-progress {
  margin-top: 16px;
  padding: 14px 16px;
  background: var(--el-fill-color-light);
  border-radius: 8px;
  border: 1px solid var(--el-border-color-lighter);
}

.foundation-progress-header {
  display: flex;
  align-items: center;
  margin-bottom: 12px;
}

.foundation-progress-title {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 14px;
  font-weight: 600;
}

.foundation-progress-title .el-icon {
  font-size: 18px;
}

.foundation-done-text {
  color: var(--el-color-success);
}

.foundation-error-text {
  color: var(--el-color-danger);
}

.foundation-stages {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.foundation-stage {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 6px 12px;
  border-radius: 20px;
  font-size: 13px;
  background: var(--el-bg-color);
  border: 1px solid var(--el-border-color);
  transition: all 0.3s;
}

.foundation-stage.pending {
  opacity: 0.5;
}

.foundation-stage.running {
  border-color: var(--el-color-primary);
  background: var(--el-color-primary-light-9);
  color: var(--el-color-primary);
  animation: foundation-pulse 1.5s ease-in-out infinite;
}

@keyframes foundation-pulse {
  0%, 100% { box-shadow: 0 0 0 0 rgba(64, 158, 255, 0.3); }
  50% { box-shadow: 0 0 0 4px rgba(64, 158, 255, 0); }
}

.foundation-stage.done {
  border-color: var(--el-color-success);
  background: var(--el-color-success-light-9);
  color: var(--el-color-success);
}

.foundation-stage.failed {
  border-color: var(--el-color-danger);
  background: var(--el-color-danger-light-9);
  color: var(--el-color-danger);
}

.foundation-stage-icon {
  font-size: 16px;
}

.foundation-stage-status {
  font-size: 12px;
  opacity: 0.8;
}

.foundation-error-detail {
  margin-top: 10px;
  padding: 8px 12px;
  background: var(--el-color-danger-light-9);
  border-radius: 4px;
  font-size: 12px;
  color: var(--el-color-danger);
  line-height: 1.5;
}
</style>

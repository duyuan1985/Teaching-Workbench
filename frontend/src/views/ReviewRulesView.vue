<template>
  <div class="page-container">
    <div class="panel">
      <div class="panel-header">
        <h2>蓝本审查规则配置</h2>
        <el-button type="primary" @click="openCreate">新增规则</el-button>
      </div>
      <p class="muted">
        为不同课程配置PPT分组方式、学习目标提取关键字、技能点关键字等规则。
        新课程加入时，只需配置一次规则，后续蓝本审查将按规则自动适配。
      </p>
    </div>

    <div class="panel">
      <h2>AI 蓝本审查</h2>
      <p class="muted">
        启用后，蓝本审查将使用智谱AI分析PPT内容，自动识别项目分组、学习目标和技能点，无需手动配置规则。
      </p>
      <div class="ai-status">
        <el-tag :type="aiEnabled ? 'success' : 'info'">
          {{ aiEnabled ? '已启用' : '未启用' }}
        </el-tag>
        <el-button :type="aiEnabled ? 'warning' : 'primary'" @click="toggleAI" :loading="togglingAI">
          {{ aiEnabled ? '关闭 AI 审查' : '启用 AI 审查' }}
        </el-button>
        <span class="muted" v-if="aiEnabled">
          启用后将优先使用AI分析，配置规则仅在AI关闭时生效
        </span>
      </div>
    </div>

    <div class="panel" v-loading="loading">
      <h2>规则列表</h2>
      <el-table :data="rules" stripe empty-text="暂无规则，点击「新增规则」创建">
        <el-table-column prop="course_name" label="课程名称" min-width="140" />
        <el-table-column prop="rule_name" label="规则名称" min-width="120" />
        <el-table-column prop="ppt_group_mode" label="PPT分组方式" min-width="120">
          <template #default="{ row }">
            <el-tag size="small">{{ groupModeText(row.ppt_group_mode) }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="objective_keyword" label="目标关键字" min-width="100" />
        <el-table-column prop="skill_keywords" label="技能关键字" min-width="140" show-overflow-tooltip />
        <el-table-column prop="is_active" label="状态" width="80">
          <template #default="{ row }">
            <el-tag :type="row.is_active ? 'success' : 'info'" size="small">
              {{ row.is_active ? '启用' : '停用' }}
            </el-tag>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="150" fixed="right">
          <template #default="{ row }">
            <el-button link type="primary" size="small" @click="openEdit(row)">编辑</el-button>
            <el-button link type="danger" size="small" @click="deleteRule(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </div>

    <el-dialog v-model="dialogVisible" :title="editing ? '编辑规则' : '新增规则'" width="640px">
      <el-form :model="form" label-width="130px">
        <el-form-item label="课程名称" required>
          <el-input v-model="form.course_name" placeholder="如：Python程序设计" />
        </el-form-item>
        <el-form-item label="规则名称">
          <el-input v-model="form.rule_name" placeholder="如：Python课程PPT分组规则" />
        </el-form-item>
        <el-form-item label="PPT分组方式">
          <el-select v-model="form.ppt_group_mode" style="width: 100%">
            <el-option label="每个PPT独立一组" value="each" />
            <el-option label="按项目目录分组" value="project_dir" />
            <el-option label="全部合并为一组" value="merge_all" />
          </el-select>
        </el-form-item>
        <el-form-item label="项目匹配正则" v-if="form.ppt_group_mode === 'project_dir'">
          <el-input v-model="form.project_pattern" placeholder="如：项目\s*0*(\d+)" />
        </el-form-item>
        <el-form-item label="大纲关键字">
          <el-input v-model="form.outline_keyword" placeholder="如：教学大纲" />
        </el-form-item>
        <el-form-item label="学习目标关键字">
          <el-input v-model="form.objective_keyword" placeholder="如：学习目标" />
        </el-form-item>
        <el-form-item label="技能关键字">
          <el-input v-model="form.skill_keywords" type="textarea" :rows="2"
            placeholder="多个关键字用英文逗号分隔，如：技能点,任务,实操" />
        </el-form-item>
        <el-form-item label="启用">
          <el-switch v-model="form.is_active" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="dialogVisible = false">取消</el-button>
        <el-button type="primary" @click="saveRule" :loading="saving">保存</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { reviewRulesApi, aiReviewApi, settingsApi } from '../api'

const loading = ref(false)
const saving = ref(false)
const togglingAI = ref(false)
const rules = ref([])
const aiEnabled = ref(false)
const dialogVisible = ref(false)
const editing = ref(false)
const form = reactive({
  course_name: '',
  rule_name: '',
  ppt_group_mode: 'each',
  project_pattern: '',
  outline_keyword: '教学大纲',
  title_extraction: 'first_slide',
  objective_keyword: '学习目标',
  skill_keywords: '',
  modernization_tags: '',
  is_active: true,
})

const groupModeText = (mode) => {
  const map = { each: '每组一个PPT', project_dir: '按项目目录', merge_all: '全部合并' }
  return map[mode] || mode
}

async function loadData() {
  loading.value = true
  try {
    const [ruleList, settings] = await Promise.all([
      reviewRulesApi.list(),
      settingsApi.get(),
    ])
    rules.value = ruleList
    aiEnabled.value = settings.ai_curriculum_review
  } finally {
    loading.value = false
  }
}

function openCreate() {
  editing.value = false
  Object.assign(form, {
    course_name: '', rule_name: '', ppt_group_mode: 'each', project_pattern: '',
    outline_keyword: '教学大纲', title_extraction: 'first_slide',
    objective_keyword: '学习目标', skill_keywords: '', modernization_tags: '', is_active: true,
  })
  dialogVisible.value = true
}

function openEdit(row) {
  editing.value = true
  Object.assign(form, {
    course_name: row.course_name, rule_name: row.rule_name,
    ppt_group_mode: row.ppt_group_mode, project_pattern: row.project_pattern || '',
    outline_keyword: row.outline_keyword || '教学大纲',
    title_extraction: row.title_extraction || 'first_slide',
    objective_keyword: row.objective_keyword || '学习目标',
    skill_keywords: row.skill_keywords || '',
    modernization_tags: row.modernization_tags || '',
    is_active: !!row.is_active,
  })
  form._id = row.id
  dialogVisible.value = true
}

async function saveRule() {
  if (!form.course_name.trim()) {
    ElMessage.warning('请输入课程名称')
    return
  }
  saving.value = true
  try {
    if (editing.value) {
      await reviewRulesApi.update(form._id, { ...form })
    } else {
      await reviewRulesApi.create({ ...form })
    }
    ElMessage.success('保存成功')
    dialogVisible.value = false
    await loadData()
  } finally {
    saving.value = false
  }
}

async function deleteRule(row) {
  try {
    await ElMessageBox.confirm(`确认删除「${row.course_name}」的审查规则？`, '删除确认', { type: 'warning' })
  } catch {
    return
  }
  try {
    await reviewRulesApi.delete(row.id)
    ElMessage.success('已删除')
    await loadData()
  } catch (e) {
    ElMessage.error('删除失败')
  }
}

async function toggleAI() {
  togglingAI.value = true
  try {
    const result = await aiReviewApi.toggle()
    aiEnabled.value = result.enabled
    ElMessage.success(result.enabled ? 'AI蓝本审查已启用' : 'AI蓝本审查已关闭')
  } finally {
    togglingAI.value = false
  }
}

onMounted(loadData)
</script>

<style scoped>
.panel-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 16px;
}
.panel-header h2 {
  margin: 0;
}
.ai-status {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-top: 12px;
}
</style>

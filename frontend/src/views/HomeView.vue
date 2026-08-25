<template>
  <div class="page-container">
    <!-- 统计卡片 -->
    <div class="stats-row">
      <el-card class="stat-card" shadow="hover" v-loading="statsLoading">
        <div class="stat-value">{{ stats.total }}</div>
        <div class="stat-label">课程实例</div>
      </el-card>
      <el-card class="stat-card" shadow="hover" v-loading="statsLoading">
        <div class="stat-value">{{ stats.terms }}</div>
        <div class="stat-label">学期数</div>
      </el-card>
      <el-card class="stat-card" shadow="hover" v-loading="statsLoading">
        <div class="stat-value">{{ stats.total_hours }}</div>
        <div class="stat-label">总学时</div>
      </el-card>
    </div>

    <!-- 课程列表 -->
    <div class="panel">
      <div class="panel-header">
        <h2>课程档案</h2>
        <div class="actions">
          <el-input
            v-model="searchQuery"
            placeholder="按课程名 / 编号 / 专业 / 班级搜索"
            clearable
            style="width: 280px"
            :prefix-icon="Search"
            @keyup.enter="loadOfferings"
            @clear="loadOfferings"
          />
          <el-select
            v-model="selectedTerm"
            placeholder="全部学期"
            clearable
            style="width: 180px"
            @change="loadOfferings"
          >
            <el-option
              v-for="term in terms"
              :key="term"
              :label="term"
              :value="term"
            />
          </el-select>
          <el-button type="primary" :icon="Plus" @click="showCreateDialog = true">
            新建课程实例
          </el-button>
        </div>
      </div>

      <el-table
        :data="offerings"
        v-loading="tableLoading"
        stripe
        empty-text="暂无课程实例，点击「新建课程实例」创建"
        style="width: 100%"
      >
        <el-table-column prop="term" label="学期" width="130" />
        <el-table-column prop="course_name" label="课程名" min-width="140" show-overflow-tooltip />
        <el-table-column prop="course_code" label="课程编号" width="120" show-overflow-tooltip />
        <el-table-column prop="major" label="专业" min-width="120" show-overflow-tooltip />
        <el-table-column prop="teaching_class" label="班级" width="130" show-overflow-tooltip>
          <template #default="{ row }">
            {{ row.teaching_class || '—' }}
          </template>
        </el-table-column>
        <el-table-column prop="total_hours" label="学时" width="80" align="center" />
        <el-table-column label="操作" width="200" fixed="right">
          <template #default="{ row }">
            <el-button size="small" type="primary" text :icon="View" @click="goDetail(row.id)">
              查看详情
            </el-button>
            <el-button size="small" type="success" text :icon="Calendar" @click="goAttendance(row.id)">
              上课
            </el-button>
          </template>
        </el-table-column>
      </el-table>
    </div>

    <!-- 新建课程实例对话框 -->
    <el-dialog
      v-model="showCreateDialog"
      title="新建课程实例"
      width="720px"
      :close-on-click-modal="false"
    >
      <el-form
        ref="createFormRef"
        :model="createForm"
        :rules="createRules"
        label-width="100px"
        label-position="right"
      >
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="学期" prop="term">
              <el-input v-model="createForm.term" placeholder="如 2025-2026-2" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="课程名" prop="course_name">
              <el-input v-model="createForm.course_name" placeholder="课程名称" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="课程编号" prop="course_code">
              <el-input v-model="createForm.course_code" placeholder="课程编号" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="专业" prop="major">
              <el-input v-model="createForm.major" placeholder="专业名称" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="班级" prop="teaching_class">
              <el-input v-model="createForm.teaching_class" placeholder="班级名称" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="课程性质" prop="course_nature">
              <el-select v-model="createForm.course_nature" placeholder="选择课程性质" clearable style="width: 100%">
                <el-option label="必修课" value="必修课" />
                <el-option label="选修课" value="选修课" />
                <el-option label="限选课" value="限选课" />
              </el-select>
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="课程类型" prop="course_type">
              <el-select v-model="createForm.course_type" placeholder="选择课程类型" clearable style="width: 100%">
                <el-option
                  v-for="t in courseTypes"
                  :key="t.id"
                  :label="t.name"
                  :value="t.name"
                />
              </el-select>
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="课程类别" prop="offering_kind">
              <el-select v-model="createForm.offering_kind" placeholder="选择课程类别" style="width: 100%">
                <el-option label="普通课程" value="普通课程" />
                <el-option label="实训课程" value="实训课程" />
              </el-select>
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="8">
            <el-form-item label="学分" prop="credits">
              <el-input-number v-model="createForm.credits" :min="0" :precision="1" :step="0.5" controls-position="right" style="width: 100%" />
            </el-form-item>
          </el-col>
          <el-col :span="8">
            <el-form-item label="总学时" prop="total_hours">
              <el-input-number v-model="createForm.total_hours" :min="0" :step="2" controls-position="right" style="width: 100%" />
            </el-form-item>
          </el-col>
          <el-col :span="8">
            <el-form-item label="周学时" prop="weekly_hours">
              <el-input-number v-model="createForm.weekly_hours" :min="0" :step="1" controls-position="right" style="width: 100%" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-row :gutter="20">
          <el-col :span="12">
            <el-form-item label="考核类型" prop="assessment_type">
              <el-input v-model="createForm.assessment_type" placeholder="如 期末考核" />
            </el-form-item>
          </el-col>
          <el-col :span="12">
            <el-form-item label="考核方式" prop="assessment_method">
              <el-input v-model="createForm.assessment_method" placeholder="如 实操" />
            </el-form-item>
          </el-col>
        </el-row>
        <el-form-item label="教材版本" prop="textbook_version">
          <el-input v-model="createForm.textbook_version" placeholder="教材版本信息" />
        </el-form-item>
        <el-form-item label="教材路径" prop="textbook_path">
          <el-input v-model="createForm.textbook_path" placeholder="教材文件路径" />
        </el-form-item>
        <el-form-item label="模板路径" prop="template_path">
          <el-input v-model="createForm.template_path" placeholder="模板文件路径" />
        </el-form-item>
        <el-form-item label="排课路径" prop="schedule_path">
          <el-input v-model="createForm.schedule_path" placeholder="排课文件路径" />
        </el-form-item>
        <el-form-item label="备注" prop="notes">
          <el-input v-model="createForm.notes" type="textarea" :rows="2" placeholder="备注信息" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showCreateDialog = false">取消</el-button>
        <el-button type="primary" :loading="createLoading" @click="handleCreate">创建</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { Search, Plus, View, Calendar } from '@element-plus/icons-vue'
import { offeringsApi, courseTypesApi } from '../api'

const router = useRouter()

// ---- 统计 ----
const stats = reactive({ total: 0, terms: 0, total_hours: 0 })
const statsLoading = ref(false)

async function loadStats() {
  statsLoading.value = true
  try {
    const data = await offeringsApi.stats()
    stats.total = data.total
    stats.terms = data.terms
    stats.total_hours = data.total_hours
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    statsLoading.value = false
  }
}

// ---- 学期列表 ----
const terms = ref([])

async function loadTerms() {
  try {
    terms.value = await offeringsApi.terms()
  } catch {
    // 错误已由 axios 拦截器提示
  }
}

// ---- 课程列表 ----
const offerings = ref([])
const tableLoading = ref(false)
const searchQuery = ref('')
const selectedTerm = ref('')

async function loadOfferings() {
  tableLoading.value = true
  try {
    const params = {}
    if (searchQuery.value) params.q = searchQuery.value
    if (selectedTerm.value) params.term = selectedTerm.value
    offerings.value = await offeringsApi.list(params)
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    tableLoading.value = false
  }
}

// ---- 跳转 ----
function goDetail(id) {
  router.push(`/offering/${id}`)
}

function goAttendance(id) {
  router.push({ path: '/attendance', query: { offering_id: id } })
}

// ---- 新建课程实例 ----
const showCreateDialog = ref(false)
const createLoading = ref(false)
const createFormRef = ref(null)

const createForm = reactive({
  term: '',
  course_name: '',
  course_code: '',
  major: '',
  teaching_class: '',
  course_nature: '',
  course_type: '',
  assessment_type: '期末考核',
  assessment_method: '实操',
  credits: 0,
  total_hours: 0,
  weekly_hours: 0,
  textbook_version: '',
  textbook_path: '',
  template_path: '',
  schedule_path: '',
  notes: '',
  offering_kind: '普通课程',
})

const createRules = {
  term: [{ required: true, message: '请输入学期', trigger: 'blur' }],
  course_name: [{ required: true, message: '请输入课程名', trigger: 'blur' }],
  major: [{ required: true, message: '请输入专业', trigger: 'blur' }],
}

function resetCreateForm() {
  Object.assign(createForm, {
    term: '',
    course_name: '',
    course_code: '',
    major: '',
    teaching_class: '',
    course_nature: '',
    course_type: '',
    assessment_type: '期末考核',
    assessment_method: '实操',
    credits: 0,
    total_hours: 0,
    weekly_hours: 0,
    textbook_version: '',
    textbook_path: '',
    template_path: '',
    schedule_path: '',
    notes: '',
    offering_kind: '普通课程',
  })
}

async function handleCreate() {
  if (!createFormRef.value) return
  try {
    await createFormRef.value.validate()
  } catch {
    return
  }
  createLoading.value = true
  try {
    await offeringsApi.create({ ...createForm })
    ElMessage.success('课程实例创建成功')
    showCreateDialog.value = false
    resetCreateForm()
    await Promise.all([loadOfferings(), loadStats(), loadTerms()])
  } catch {
    // 错误已由 axios 拦截器提示
  } finally {
    createLoading.value = false
  }
}

// ---- 课程类型选项（基础设置 → 课程类型管理） ----
const courseTypes = ref([])

async function loadCourseTypes() {
  try {
    courseTypes.value = await courseTypesApi.list()
  } catch {
    // 错误已由 axios 拦截器提示
  }
}

onMounted(() => {
  loadStats()
  loadTerms()
  loadOfferings()
  loadCourseTypes()
})
</script>

<style scoped>
.stats-row {
  display: flex;
  gap: 20px;
  margin-bottom: 20px;
}

.stat-card {
  flex: 1;
  text-align: center;
}

.stat-card :deep(.el-card__body) {
  padding: 24px 20px;
}

.stat-value {
  font-size: 36px;
  font-weight: bold;
  color: var(--primary);
  line-height: 1.2;
}

.stat-label {
  font-size: 14px;
  color: var(--muted);
  margin-top: 8px;
}

.panel-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 16px;
  flex-wrap: wrap;
  gap: 12px;
}

.panel-header h2 {
  margin: 0;
}
</style>

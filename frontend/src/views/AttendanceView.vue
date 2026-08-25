<template>
  <div class="page-container">
    <!-- 顶部筛选栏 -->
    <div class="panel">
      <h2>上课点名</h2>
      <div class="filter-bar">
        <el-select
          v-model="selectedOfferingId"
          placeholder="选择课程实例"
          filterable
          clearable
          class="filter-item"
          @change="onOfferingChange"
        >
          <el-option
            v-for="o in offerings"
            :key="o.id"
            :label="offeringLabel(o)"
            :value="o.id"
          />
        </el-select>

        <el-select
          v-model="selectedClass"
          placeholder="选择班级"
          clearable
          class="filter-item"
          :disabled="!selectedOfferingId"
          @change="onClassChange"
        >
          <el-option
            v-for="c in classes"
            :key="c"
            :label="c"
            :value="c"
          />
        </el-select>

        <el-select
          v-model="selectedDate"
          placeholder="选择上课日期"
          clearable
          class="filter-item"
          :disabled="!selectedOfferingId"
          @change="onDateChange"
        >
          <el-option
            v-for="s in sessions"
            :key="s.lesson_date"
            :label="sessionLabel(s)"
            :value="s.lesson_date"
          />
        </el-select>
      </div>
      <p v-if="!selectedOfferingId" class="muted">请先选择课程实例，再选择班级和上课日期。</p>
    </div>

    <!-- 学生考勤表格 -->
    <div class="panel" v-loading="loadingAttendance">
      <h2>
        学生考勤
        <span v-if="selectedClass" class="badge">{{ selectedClass }}</span>
        <span v-if="selectedDate" class="badge">{{ selectedDate }}</span>
      </h2>
      <el-table :data="students" border stripe style="width: 100%" empty-text="暂无学生数据">
        <el-table-column label="序号" type="index" width="60" align="center" />
        <el-table-column prop="student_no" label="学号" width="130" />
        <el-table-column prop="student_name" label="姓名" width="120" />
        <el-table-column prop="gender" label="性别" width="70" align="center" />
        <el-table-column label="考勤状态" width="150">
          <template #default="{ row }">
            <el-select
              :model-value="getStatus(row.id)"
              size="small"
              @update:model-value="(val) => onAttendanceChange(row.id, 'status', val)"
            >
              <el-option
                v-for="opt in statusOptions"
                :key="opt.value"
                :label="opt.label"
                :value="opt.value"
              />
            </el-select>
          </template>
        </el-table-column>
        <el-table-column label="备注" min-width="200">
          <template #default="{ row }">
            <el-input
              :model-value="getNotes(row.id)"
              size="small"
              placeholder="备注信息"
              @update:model-value="(val) => onAttendanceChange(row.id, 'notes', val)"
            />
          </template>
        </el-table-column>
      </el-table>
      <div class="actions" style="margin-top: 16px">
        <el-button
          type="primary"
          :loading="saving"
          :disabled="!canSave"
          @click="saveAttendance"
        >
          保存点名
        </el-button>
        <el-button
          v-if="canSave"
          @click="markAllPresent"
        >
          全部出勤
        </el-button>
      </div>
    </div>

    <!-- 考勤规则 + 成绩构成 -->
    <div class="dual-panels" v-if="selectedOfferingId">
      <!-- 考勤规则 -->
      <div class="panel">
        <h2>考勤规则</h2>
        <el-table :data="rules" border size="small" empty-text="暂无规则">
          <el-table-column prop="status" label="状态" width="120" />
          <el-table-column label="扣分" width="90">
            <template #default="{ row }">{{ row.deduction }}</template>
          </el-table-column>
          <el-table-column prop="sort_order" label="排序" width="70" />
          <el-table-column label="操作" width="80">
            <template #default="{ row }">
              <el-button type="danger" link size="small" @click="deleteRule(row)">删除</el-button>
            </template>
          </el-table-column>
        </el-table>
        <el-form :model="ruleForm" inline size="small" class="add-form">
          <el-form-item label="状态">
            <el-select v-model="ruleForm.status" placeholder="选择状态" style="width: 120px">
              <el-option
                v-for="opt in statusOptions"
                :key="opt.value"
                :label="opt.label"
                :value="opt.value"
              />
            </el-select>
          </el-form-item>
          <el-form-item label="扣分">
            <el-input-number v-model="ruleForm.deduction" :min="0" :step="0.5" controls-position="right" style="width: 100px" />
          </el-form-item>
          <el-form-item label="排序">
            <el-input-number v-model="ruleForm.sort_order" :min="0" :step="1" controls-position="right" style="width: 90px" />
          </el-form-item>
          <el-form-item>
            <el-button type="primary" size="small" @click="addRule">添加规则</el-button>
          </el-form-item>
        </el-form>
      </div>

      <!-- 成绩构成 -->
      <div class="panel">
        <h2>成绩构成</h2>
        <el-table :data="components" border size="small" empty-text="暂无构成项">
          <el-table-column prop="component_name" label="名称" min-width="120" />
          <el-table-column label="权重" width="80">
            <template #default="{ row }">{{ row.weight }}</template>
          </el-table-column>
          <el-table-column prop="source_type" label="来源" width="90" />
          <el-table-column prop="sort_order" label="排序" width="70" />
          <el-table-column label="操作" width="80">
            <template #default="{ row }">
              <el-button type="danger" link size="small" @click="deleteComponent(row)">删除</el-button>
            </template>
          </el-table-column>
        </el-table>
        <el-form :model="componentForm" inline size="small" class="add-form">
          <el-form-item label="名称">
            <el-input v-model="componentForm.component_name" placeholder="构成名称" style="width: 140px" />
          </el-form-item>
          <el-form-item label="权重">
            <el-input-number v-model="componentForm.weight" :min="0" :step="1" controls-position="right" style="width: 100px" />
          </el-form-item>
          <el-form-item label="来源">
            <el-select v-model="componentForm.source_type" placeholder="来源" style="width: 120px" filterable allow-create default-first-option>
              <el-option label="手工" value="手工" />
              <el-option label="考勤" value="考勤" />
              <el-option label="作业" value="作业" />
              <el-option label="期中" value="期中" />
              <el-option label="期末" value="期末" />
            </el-select>
          </el-form-item>
          <el-form-item label="排序">
            <el-input-number v-model="componentForm.sort_order" :min="0" :step="1" controls-position="right" style="width: 90px" />
          </el-form-item>
          <el-form-item>
            <el-button type="primary" size="small" @click="addComponent">添加构成</el-button>
          </el-form-item>
        </el-form>
      </div>
    </div>

    <!-- 请假记录 -->
    <div class="panel" v-if="selectedOfferingId">
      <h2>请假记录</h2>
      <el-table :data="leaves" border size="small" empty-text="暂无请假记录">
        <el-table-column prop="student_no" label="学号" width="130" />
        <el-table-column prop="student_name" label="姓名" width="120" />
        <el-table-column prop="start_date" label="开始日期" width="120" />
        <el-table-column prop="end_date" label="结束日期" width="120" />
        <el-table-column prop="reason" label="事由" min-width="200" />
        <el-table-column label="操作" width="80">
          <template #default="{ row }">
            <el-button type="danger" link size="small" @click="deleteLeave(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
      <el-form :model="leaveForm" inline size="small" class="add-form">
        <el-form-item label="学生">
          <el-select v-model="leaveForm.student_id" placeholder="选择学生" filterable style="width: 180px">
            <el-option
              v-for="s in students"
              :key="s.id"
              :label="`${s.student_no} ${s.student_name}`"
              :value="s.id"
            />
          </el-select>
        </el-form-item>
        <el-form-item label="开始">
          <el-date-picker v-model="leaveForm.start_date" type="date" value-format="YYYY-MM-DD" placeholder="开始日期" style="width: 150px" />
        </el-form-item>
        <el-form-item label="结束">
          <el-date-picker v-model="leaveForm.end_date" type="date" value-format="YYYY-MM-DD" placeholder="结束日期" style="width: 150px" />
        </el-form-item>
        <el-form-item label="事由">
          <el-input v-model="leaveForm.reason" placeholder="请假事由" style="width: 200px" />
        </el-form-item>
        <el-form-item>
          <el-button type="primary" size="small" @click="addLeave">添加请假</el-button>
        </el-form-item>
      </el-form>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, reactive } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import {
  attendanceApi,
  attendanceRulesApi,
  gradeComponentsApi,
  leavesApi,
  offeringsApi,
} from '../api'

// ---- 常量 ----
const statusOptions = [
  { label: '出勤', value: '出勤' },
  { label: '迟到', value: '迟到' },
  { label: '早退', value: '早退' },
  { label: '请假', value: '请假' },
  { label: '旷课', value: '旷课' },
]

// ---- 筛选状态 ----
const offerings = ref([])
const selectedOfferingId = ref(null)
const classes = ref([])
const selectedClass = ref('')
const sessions = ref([])
const selectedDate = ref('')

// ---- 数据 ----
const students = ref([])
const records = ref({})
const rules = ref([])
const components = ref([])
const leaves = ref([])

// ---- 考勤编辑表单 ----
const attendanceMap = ref({})

// ---- 加载状态 ----
const loadingAttendance = ref(false)
const saving = ref(false)

// ---- 规则添加表单 ----
const ruleForm = reactive({
  status: '出勤',
  deduction: 0,
  sort_order: 0,
})

// ---- 成绩构成添加表单 ----
const componentForm = reactive({
  component_name: '',
  weight: 0,
  source_type: '手工',
  sort_order: 0,
})

// ---- 请假添加表单 ----
const leaveForm = reactive({
  student_id: null,
  start_date: '',
  end_date: '',
  reason: '',
})

// ---- 计算属性 ----
const canSave = computed(() => selectedOfferingId.value && selectedClass.value && selectedDate.value && students.value.length > 0)

// ---- 工具函数 ----
function offeringLabel(o) {
  const parts = [o.course_name, o.term]
  if (o.major) parts.push(o.major)
  return parts.join(' · ')
}

function sessionLabel(s) {
  let label = s.lesson_date
  if (s.week_no) label += ` (第${s.week_no}周)`
  if (s.periods) label += ` ${s.periods}`
  if (s.classroom) label += ` ${s.classroom}`
  return label
}

// ---- 初始化考勤映射 ----
function initAttendanceMap() {
  const map = {}
  for (const s of students.value) {
    const key = String(s.id)
    const rec = records.value[key]
    map[key] = {
      status: rec?.status || '出勤',
      notes: rec?.notes || '',
    }
  }
  attendanceMap.value = map
}

// ---- 加载考勤数据 ----
async function loadAttendance() {
  if (!selectedOfferingId.value) return
  loadingAttendance.value = true
  try {
    const res = await attendanceApi.get(
      selectedOfferingId.value,
      selectedClass.value,
      selectedDate.value,
    )
    classes.value = res.classes || []
    if (res.selected_class) selectedClass.value = res.selected_class
    sessions.value = res.sessions || []
    students.value = res.students || []
    records.value = res.records || {}
    rules.value = res.rules || []
    components.value = res.components || []
    leaves.value = res.leaves || []
    initAttendanceMap()
  } catch (e) {
    // 错误已由拦截器提示
  } finally {
    loadingAttendance.value = false
  }
}

// ---- 筛选事件 ----
async function onOfferingChange() {
  selectedClass.value = ''
  selectedDate.value = ''
  if (!selectedOfferingId.value) {
    classes.value = []
    sessions.value = []
    students.value = []
    rules.value = []
    components.value = []
    leaves.value = []
    attendanceMap.value = {}
    return
  }
  await loadAttendance()
}

async function onClassChange() {
  await loadAttendance()
}

async function onDateChange() {
  await loadAttendance()
}

// ---- 考勤表格内联编辑 ----
function getStatus(studentId) {
  const entry = attendanceMap.value[String(studentId)]
  return entry?.status || '出勤'
}

function getNotes(studentId) {
  const entry = attendanceMap.value[String(studentId)]
  return entry?.notes || ''
}

function onAttendanceChange(studentId, field, val) {
  const key = String(studentId)
  if (!attendanceMap.value[key]) {
    attendanceMap.value[key] = { status: '出勤', notes: '' }
  }
  attendanceMap.value[key][field] = val
}

// ---- 全部出勤 ----
function markAllPresent() {
  for (const s of students.value) {
    const key = String(s.id)
    if (attendanceMap.value[key]) {
      attendanceMap.value[key].status = '出勤'
    }
  }
}

// ---- 保存点名 ----
async function saveAttendance() {
  if (!canSave.value) return
  saving.value = true
  try {
    const studentIds = students.value.map((s) => s.id)
    const statuses = {}
    const notes = {}
    const scores = {}
    for (const s of students.value) {
      const key = String(s.id)
      const entry = attendanceMap.value[key] || {}
      statuses[key] = entry.status || '出勤'
      notes[key] = entry.notes || ''
      scores[key] = entry.score ?? 0
    }
    await attendanceApi.save(
      selectedOfferingId.value,
      selectedClass.value,
      selectedDate.value,
      { student_ids: studentIds, statuses, notes, scores },
    )
    ElMessage.success('点名保存成功')
    await loadAttendance()
  } catch (e) {
    // 错误已由拦截器提示
  } finally {
    saving.value = false
  }
}

// ---- 考勤规则 ----
async function addRule() {
  if (!ruleForm.status) {
    ElMessage.warning('请选择考勤状态')
    return
  }
  try {
    await attendanceRulesApi.create({
      offering_id: selectedOfferingId.value,
      status: ruleForm.status,
      deduction: ruleForm.deduction,
      sort_order: ruleForm.sort_order,
    })
    ElMessage.success('规则已添加')
    ruleForm.status = '出勤'
    ruleForm.deduction = 0
    ruleForm.sort_order = 0
    await loadAttendance()
  } catch (e) {
    // 已提示
  }
}

async function deleteRule(row) {
  try {
    await ElMessageBox.confirm(`确定删除规则"${row.status}"？`, '提示', { type: 'warning' })
    await attendanceRulesApi.delete(row.id)
    ElMessage.success('规则已删除')
    await loadAttendance()
  } catch (e) {
    // 取消或已提示
  }
}

// ---- 成绩构成 ----
async function addComponent() {
  if (!componentForm.component_name) {
    ElMessage.warning('请输入构成名称')
    return
  }
  try {
    await gradeComponentsApi.create({
      offering_id: selectedOfferingId.value,
      component_name: componentForm.component_name,
      weight: componentForm.weight,
      source_type: componentForm.source_type,
      sort_order: componentForm.sort_order,
    })
    ElMessage.success('构成项已添加')
    componentForm.component_name = ''
    componentForm.weight = 0
    componentForm.source_type = '手工'
    componentForm.sort_order = 0
    await loadAttendance()
  } catch (e) {
    // 已提示
  }
}

async function deleteComponent(row) {
  try {
    await ElMessageBox.confirm(`确定删除构成项"${row.component_name}"？`, '提示', { type: 'warning' })
    await gradeComponentsApi.delete(row.id)
    ElMessage.success('构成项已删除')
    await loadAttendance()
  } catch (e) {
    // 取消或已提示
  }
}

// ---- 请假记录 ----
async function addLeave() {
  if (!leaveForm.student_id) {
    ElMessage.warning('请选择学生')
    return
  }
  if (!leaveForm.start_date || !leaveForm.end_date) {
    ElMessage.warning('请选择请假日期')
    return
  }
  try {
    await leavesApi.create({
      offering_id: selectedOfferingId.value,
      student_id: leaveForm.student_id,
      start_date: leaveForm.start_date,
      end_date: leaveForm.end_date,
      reason: leaveForm.reason,
    })
    ElMessage.success('请假记录已添加')
    leaveForm.student_id = null
    leaveForm.start_date = ''
    leaveForm.end_date = ''
    leaveForm.reason = ''
    await loadAttendance()
  } catch (e) {
    // 已提示
  }
}

async function deleteLeave(row) {
  try {
    await ElMessageBox.confirm(`确定删除${row.student_name}的请假记录？`, '提示', { type: 'warning' })
    await leavesApi.delete(row.id)
    ElMessage.success('请假记录已删除')
    await loadAttendance()
  } catch (e) {
    // 取消或已提示
  }
}

// ---- 页面初始化 ----
onMounted(async () => {
  try {
    const res = await offeringsApi.list()
    offerings.value = Array.isArray(res) ? res : []
  } catch (e) {
    offerings.value = []
  }
})
</script>

<style scoped>
.filter-bar {
  display: flex;
  gap: 12px;
  flex-wrap: wrap;
  align-items: center;
}

.filter-item {
  width: 280px;
}

.dual-panels {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 20px;
}

@media (max-width: 1100px) {
  .dual-panels {
    grid-template-columns: 1fr;
  }
}

.add-form {
  margin-top: 12px;
  padding-top: 12px;
  border-top: 1px dashed var(--border);
}
</style>

<template>
  <div class="page-container">
    <!-- 作业信息 -->
    <div class="panel" v-loading="loading">
      <div class="actions mb16">
        <el-button :icon="ArrowLeft" @click="$router.back()">返回</el-button>
        <h2 style="margin-bottom: 0"><el-icon><Document /></el-icon> 作业成绩</h2>
      </div>
      <el-descriptions v-if="assignment" :column="3" border size="small">
        <el-descriptions-item label="课程">{{ assignment.course_name || '-' }}</el-descriptions-item>
        <el-descriptions-item label="作业名称">{{ assignment.name || assignment.title || '-' }}</el-descriptions-item>
        <el-descriptions-item label="满分">{{ assignment.max_score ?? '-' }}</el-descriptions-item>
      </el-descriptions>

      <div class="actions mt16">
        <el-form-item label="班级" style="margin-bottom: 0">
          <el-select v-model="selectedClass" placeholder="选择班级" filterable @change="loadData" style="width: 200px">
            <el-option v-for="c in classes" :key="c" :label="c" :value="c" />
          </el-select>
        </el-form-item>
        <el-upload
          ref="uploadRef"
          :auto-upload="false"
          :show-file-list="false"
          accept=".xlsx,.xls"
          :on-change="handleExcelChange"
        >
          <el-button type="success" :icon="Upload" :loading="importing">导入Excel成绩</el-button>
        </el-upload>
        <el-button type="primary" :icon="Check" :loading="saving" @click="handleSave">保存成绩</el-button>
      </div>
    </div>

    <!-- 学生成绩表格 -->
    <div class="panel">
      <h2><el-icon><Tickets /></el-icon> 学生成绩</h2>
      <el-table :data="tableData" v-loading="loading" border stripe empty-text="暂无数据" max-height="500">
        <el-table-column type="index" label="序号" width="60" align="center" />
        <el-table-column prop="student_id" label="学号" width="140" />
        <el-table-column prop="name" label="姓名" width="120" />
        <el-table-column label="得分" width="140" align="center">
          <template #default="{ row }">
            <el-input-number
              v-model="row.score"
              :min="0"
              :max="assignment?.max_score || 100"
              :precision="1"
              :controls="false"
              size="small"
              style="width: 80px"
            />
          </template>
        </el-table-column>
        <el-table-column label="来源" width="140" align="center">
          <template #default="{ row }">
            <el-tag size="small" :type="row.source === '导入' ? 'success' : row.source ? 'info' : 'plain'">
              {{ row.source || '-' }}
            </el-tag>
          </template>
        </el-table-column>
      </el-table>
    </div>

    <!-- 导入异常记录 -->
    <div class="panel" v-if="issues.length">
      <h2><el-icon><WarningFilled /></el-icon> 导入异常记录</h2>
      <el-table :data="issues" border stripe max-height="300">
        <el-table-column type="index" label="序号" width="60" align="center" />
        <el-table-column prop="row" label="Excel行" width="80" align="center" />
        <el-table-column prop="student_id" label="学号" width="140" />
        <el-table-column prop="message" label="异常说明" min-width="300" show-overflow-tooltip />
      </el-table>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted } from 'vue'
import { useRoute } from 'vue-router'
import { ElMessage } from 'element-plus'
import { ArrowLeft, Upload, Check, Document, Tickets, WarningFilled } from '@element-plus/icons-vue'
import { assignmentsApi } from '../api'
import * as XLSX from 'xlsx'

const route = useRoute()
const assignmentId = computed(() => route.params.id)

const loading = ref(false)
const importing = ref(false)
const saving = ref(false)

const assignment = ref(null)
const classes = ref([])
const selectedClass = ref('')
const students = ref([])
const scoresMap = ref({})
const issues = ref([])

const uploadRef = ref(null)

const tableData = computed(() => {
  return students.value.map((s) => {
    const sc = scoresMap.value[s.id] || scoresMap.value[s.student_id] || {}
    return {
      student_id: s.student_id ?? s.id,
      name: s.name,
      score: sc.score ?? null,
      source: sc.source || '',
    }
  })
})

async function loadData() {
  if (!selectedClass.value) return
  loading.value = true
  issues.value = []
  try {
    const data = await assignmentsApi.get(assignmentId.value, selectedClass.value)
    assignment.value = data.assignment
    classes.value = data.classes || []
    students.value = data.students || []
    scoresMap.value = data.scores || {}
    if (data.issues) issues.value = data.issues
  } finally {
    loading.value = false
  }
}

async function loadInitial() {
  loading.value = true
  try {
    const data = await assignmentsApi.get(assignmentId.value, '')
    assignment.value = data.assignment
    classes.value = data.classes || []
    selectedClass.value = data.selected_class || (classes.value.length ? classes.value[0] : '')
    students.value = data.students || []
    scoresMap.value = data.scores || {}
    if (data.issues) issues.value = data.issues
  } finally {
    loading.value = false
  }
}

async function handleExcelChange(file) {
  if (!selectedClass.value) {
    ElMessage.warning('请先选择班级')
    uploadRef.value?.clearFiles()
    return
  }
  importing.value = true
  try {
    const data = await file.raw.arrayBuffer()
    const workbook = XLSX.read(data, { type: 'array' })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' })

    const newScores = {}
    const newIssues = []
    const studentMap = {}

    for (const s of students.value) {
      const sid = s.student_id ?? s.id
      if (sid) studentMap[String(sid)] = s
    }

    rows.forEach((row, idx) => {
      const rawId = row['学号'] ?? row['student_id'] ?? row['Student ID'] ?? ''
      const rawScore = row['得分'] ?? row['score'] ?? row['Score'] ?? ''
      const sid = String(rawId).trim()
      const scoreNum = parseFloat(rawScore)

      if (!sid) {
        newIssues.push({ row: idx + 2, student_id: '', message: '学号为空' })
        return
      }
      if (!studentMap[sid]) {
        newIssues.push({ row: idx + 2, student_id: sid, message: '学号不在班级名单中' })
        return
      }
      if (isNaN(scoreNum)) {
        newIssues.push({ row: idx + 2, student_id: sid, message: `得分格式无效: "${rawScore}"` })
        return
      }
      newScores[sid] = { score: scoreNum, source: '导入' }
    })

    scoresMap.value = { ...scoresMap.value, ...newScores }
    issues.value = newIssues

    const matched = Object.keys(newScores).length
    if (newIssues.length === 0) {
      ElMessage.success(`导入成功，共 ${matched} 条记录`)
    } else {
      ElMessage.warning(`导入完成：成功 ${matched} 条，异常 ${newIssues.length} 条`)
    }
  } catch (err) {
    ElMessage.error('Excel解析失败: ' + (err.message || '未知错误'))
  } finally {
    importing.value = false
    uploadRef.value?.clearFiles()
  }
}

async function handleSave() {
  if (!selectedClass.value) {
    ElMessage.warning('请先选择班级')
    return
  }
  saving.value = true
  try {
    const payload = {}
    for (const row of tableData.value) {
      if (row.score !== null && row.score !== undefined && row.score !== '') {
        payload[row.student_id] = { score: row.score, source: row.source || '手动录入' }
      }
    }
    await assignmentsApi.saveScores(assignmentId.value, selectedClass.value, { scores: payload })
    ElMessage.success('成绩已保存')
    await loadData()
  } finally {
    saving.value = false
  }
}

onMounted(() => {
  loadInitial()
})
</script>

<style scoped>
.mb16 {
  margin-bottom: 16px;
}

.mt16 {
  margin-top: 16px;
}
</style>

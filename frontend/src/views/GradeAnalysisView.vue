<template>
  <div class="page-container">
    <el-row :gutter="20">
      <!-- 单份生成 -->
      <el-col :span="12">
        <div class="panel">
          <h2><el-icon><Document /></el-icon> 单份生成</h2>
          <el-form :model="singleForm" label-width="100px" v-loading="generating">
            <el-form-item label="课程">
              <el-select v-model="singleForm.offering_id" placeholder="选择课程" filterable style="width: 100%">
                <el-option
                  v-for="o in offerings"
                  :key="o.id"
                  :label="`${o.term} · ${o.course_name} · ${o.class_name || ''}`"
                  :value="o.id"
                />
              </el-select>
            </el-form-item>
            <el-form-item label="PDF文件">
              <el-upload
                ref="singleUploadRef"
                :auto-upload="false"
                :limit="1"
                accept=".pdf"
                :on-change="(file) => singleFile = file.raw"
                :on-remove="() => singleFile = null"
              >
                <el-button :icon="Upload">选择PDF文件</el-button>
              </el-upload>
            </el-form-item>
            <el-form-item label="考试日期">
              <el-date-picker v-model="singleForm.exam_date" type="date" value-format="YYYY-MM-DD" placeholder="选择日期" style="width: 100%" />
            </el-form-item>
            <el-form-item label="出题方式">
              <el-select v-model="singleForm.question_source" placeholder="选择出题方式" style="width: 100%">
                <el-option label="题库组卷" value="题库组卷" />
                <el-option label="教师自编" value="教师自编" />
                <el-option label="教研室统编" value="教研室统编" />
              </el-select>
            </el-form-item>
            <el-form-item label="考试方式">
              <el-select v-model="singleForm.exam_mode" placeholder="选择考试方式" style="width: 100%">
                <el-option label="闭卷" value="闭卷" />
                <el-option label="开卷" value="开卷" />
                <el-option label="机考" value="机考" />
              </el-select>
            </el-form-item>
            <el-form-item label="阅卷方式">
              <el-select v-model="singleForm.marking_mode" placeholder="选择阅卷方式" style="width: 100%">
                <el-option label="流水阅卷" value="流水阅卷" />
                <el-option label="任课教师阅卷" value="任课教师阅卷" />
              </el-select>
            </el-form-item>
            <el-form-item>
              <el-button type="primary" :loading="generating" @click="handleGenerateSingle">生成</el-button>
            </el-form-item>
          </el-form>
        </div>
      </el-col>

      <!-- 批量生成 -->
      <el-col :span="12">
        <div class="panel">
          <h2><el-icon><Files /></el-icon> 批量生成</h2>
          <el-form :model="batchForm" label-width="100px" v-loading="batchGenerating">
            <el-form-item label="课程">
              <el-select v-model="batchForm.offering_id" placeholder="选择课程" filterable style="width: 100%">
                <el-option
                  v-for="o in offerings"
                  :key="o.id"
                  :label="`${o.term} · ${o.course_name} · ${o.class_name || ''}`"
                  :value="o.id"
                />
              </el-select>
            </el-form-item>
            <el-form-item label="考试日期">
              <el-date-picker v-model="batchForm.exam_date" type="date" value-format="YYYY-MM-DD" placeholder="选择日期" style="width: 100%" />
            </el-form-item>
            <el-form-item label="出题方式">
              <el-select v-model="batchForm.question_source" placeholder="选择出题方式" style="width: 100%">
                <el-option label="题库组卷" value="题库组卷" />
                <el-option label="教师自编" value="教师自编" />
                <el-option label="教研室统编" value="教研室统编" />
              </el-select>
            </el-form-item>
            <el-form-item label="考试方式">
              <el-select v-model="batchForm.exam_mode" placeholder="选择考试方式" style="width: 100%">
                <el-option label="闭卷" value="闭卷" />
                <el-option label="开卷" value="开卷" />
                <el-option label="机考" value="机考" />
              </el-select>
            </el-form-item>
            <el-form-item label="阅卷方式">
              <el-select v-model="batchForm.marking_mode" placeholder="选择阅卷方式" style="width: 100%">
                <el-option label="流水阅卷" value="流水阅卷" />
                <el-option label="任课教师阅卷" value="任课教师阅卷" />
              </el-select>
            </el-form-item>
            <el-form-item label="PDF文件">
              <el-upload
                ref="batchUploadRef"
                :auto-upload="false"
                multiple
                accept=".pdf"
                :on-change="handleBatchFileChange"
                :on-remove="handleBatchFileRemove"
              >
                <el-button :icon="Upload">选择多个PDF文件</el-button>
                <template #tip>
                  <div class="muted">支持同时上传多个PDF文件，将逐个生成</div>
                </template>
              </el-upload>
            </el-form-item>
            <el-form-item>
              <el-button type="primary" :loading="batchGenerating" @click="handleGenerateBatch">批量生成</el-button>
            </el-form-item>
          </el-form>
        </div>
      </el-col>
    </el-row>

    <!-- 文档列表 -->
    <div class="panel">
      <div class="actions mb16">
        <h2 style="margin-bottom: 0"><el-icon><FolderOpened /></el-icon> 成绩分析文档</h2>
        <el-button :loading="loading" :icon="Refresh" @click="loadList">刷新</el-button>
      </div>
      <el-table :data="list" v-loading="loading" border stripe empty-text="暂无文档">
        <el-table-column prop="term" label="学期" width="120" />
        <el-table-column prop="course_name" label="课程" min-width="140" />
        <el-table-column prop="class_name" label="班级" width="120" />
        <el-table-column prop="source_file" label="来源文件名" min-width="180" show-overflow-tooltip />
        <el-table-column prop="created_at" label="生成时间" width="170" />
        <el-table-column label="操作" width="160" align="center" fixed="right">
          <template #default="{ row }">
            <el-button size="small" :icon="FolderOpened" @click="handleOpen(row)">打开文件夹</el-button>
            <el-button size="small" type="danger" :icon="Delete" @click="handleDelete(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </div>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { Upload, Refresh, Delete, FolderOpened, Document, Files } from '@element-plus/icons-vue'
import { gradeAnalysisApi, offeringsApi, utilApi } from '../api'

const loading = ref(false)
const generating = ref(false)
const batchGenerating = ref(false)
const offerings = ref([])
const list = ref([])

const singleFile = ref(null)
const singleUploadRef = ref(null)
const batchUploadRef = ref(null)
const batchFiles = ref([])

const singleForm = reactive({
  offering_id: null,
  exam_date: '',
  question_source: '',
  exam_mode: '',
  marking_mode: '',
})

const batchForm = reactive({
  offering_id: null,
  exam_date: '',
  question_source: '',
  exam_mode: '',
  marking_mode: '',
})

async function loadOfferings() {
  offerings.value = await offeringsApi.list()
}

async function loadList() {
  loading.value = true
  try {
    list.value = await gradeAnalysisApi.list()
  } finally {
    loading.value = false
  }
}

async function handleGenerateSingle() {
  if (!singleForm.offering_id) {
    ElMessage.warning('请选择课程')
    return
  }
  if (!singleFile.value) {
    ElMessage.warning('请选择PDF文件')
    return
  }
  generating.value = true
  try {
    const formData = new FormData()
    formData.append('offering_id', singleForm.offering_id)
    formData.append('grade_pdf', singleFile.value)
    formData.append('exam_date', singleForm.exam_date || '')
    formData.append('question_source', singleForm.question_source || '')
    formData.append('exam_mode', singleForm.exam_mode || '')
    formData.append('marking_mode', singleForm.marking_mode || '')
    await gradeAnalysisApi.generate(formData)
    ElMessage.success('生成成功')
    singleFile.value = null
    singleUploadRef.value?.clearFiles()
    await loadList()
  } finally {
    generating.value = false
  }
}

function handleBatchFileChange(file) {
  batchFiles.value.push(file.raw)
}

function handleBatchFileRemove(file) {
  const idx = batchFiles.value.indexOf(file.raw)
  if (idx > -1) batchFiles.value.splice(idx, 1)
}

async function handleGenerateBatch() {
  if (!batchForm.offering_id) {
    ElMessage.warning('请选择课程')
    return
  }
  if (!batchFiles.value.length) {
    ElMessage.warning('请选择PDF文件')
    return
  }
  batchGenerating.value = true
  let successCount = 0
  let failCount = 0
  try {
    for (const file of batchFiles.value) {
      const formData = new FormData()
      formData.append('offering_id', batchForm.offering_id)
      formData.append('grade_pdf', file)
      formData.append('exam_date', batchForm.exam_date || '')
      formData.append('question_source', batchForm.question_source || '')
      formData.append('exam_mode', batchForm.exam_mode || '')
      formData.append('marking_mode', batchForm.marking_mode || '')
      try {
        await gradeAnalysisApi.generate(formData)
        successCount++
      } catch {
        failCount++
      }
    }
    if (successCount) ElMessage.success(`成功生成 ${successCount} 份${failCount ? `，失败 ${failCount} 份` : ''}`)
    else if (failCount) ElMessage.error(`全部生成失败（${failCount} 份）`)
    batchFiles.value = []
    batchUploadRef.value?.clearFiles()
    await loadList()
  } finally {
    batchGenerating.value = false
  }
}

async function handleOpen(row) {
  await utilApi.openLocation({ offering_id: row.offering_id, kind: 'grade_analysis', document_id: row.id })
}

async function handleDelete(row) {
  try {
    await ElMessageBox.confirm(`确定删除「${row.source_file || row.course_name}」的成绩分析文档吗？`, '删除确认', {
      type: 'warning',
      confirmButtonText: '确定删除',
      cancelButtonText: '取消',
    })
  } catch {
    return
  }
  try {
    await gradeAnalysisApi.delete(row.id)
    ElMessage.success('删除成功')
    await loadList()
  } catch (e) {
    ElMessage.error('删除失败')
  }
}

onMounted(() => {
  loadOfferings()
  loadList()
})
</script>

<style scoped>
.mb16 {
  margin-bottom: 16px;
}
</style>

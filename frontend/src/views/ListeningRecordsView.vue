<template>
  <div class="page-container">
    <!-- 生成面板 -->
    <div class="panel">
      <h2><el-icon><Microphone /></el-icon> 生成听课记录</h2>
      <el-form :model="form" label-width="100px" v-loading="generating" style="max-width: 600px">
        <el-form-item label="听课模板">
          <el-select v-model="form.template_path" placeholder="选择模板" filterable style="width: 100%">
            <el-option
              v-for="t in templates"
              :key="t.path || t.id"
              :label="t.name || t.path"
              :value="t.path || t.name"
            />
          </el-select>
        </el-form-item>
        <el-form-item label="课次">
          <el-select v-model="form.session_id" placeholder="选择课次" filterable style="width: 100%">
            <el-option
              v-for="s in sessions"
              :key="s.id"
              :label="formatSessionLabel(s)"
              :value="s.id"
            />
          </el-select>
        </el-form-item>
        <el-form-item>
          <el-button type="primary" :loading="generating" :icon="Microphone" @click="handleGenerate">生成</el-button>
        </el-form-item>
      </el-form>
    </div>

    <!-- 记录列表 -->
    <div class="panel">
      <div class="actions mb16">
        <h2 style="margin-bottom: 0"><el-icon><FolderOpened /></el-icon> 听课记录文档</h2>
        <el-button :loading="loading" :icon="Refresh" @click="loadAll">刷新</el-button>
      </div>
      <el-table :data="list" v-loading="loading" border stripe empty-text="暂无文档">
        <el-table-column prop="session_date" label="日期" width="120" />
        <el-table-column prop="course_name" label="课程" min-width="140" />
        <el-table-column prop="class_name" label="班级" width="120" />
        <el-table-column prop="template_name" label="模板" min-width="180" show-overflow-tooltip />
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
import { Refresh, Delete, FolderOpened, Microphone } from '@element-plus/icons-vue'
import { listeningApi, utilApi } from '../api'

const loading = ref(false)
const generating = ref(false)
const sessions = ref([])
const templates = ref([])
const list = ref([])

const form = reactive({
  session_id: null,
  template_path: '',
})

function formatSessionLabel(s) {
  const parts = []
  if (s.date) parts.push(s.date)
  if (s.course_name) parts.push(s.course_name)
  if (s.class_name) parts.push(s.class_name)
  return parts.join(' · ')
}

async function loadAll() {
  loading.value = true
  try {
    const [sess, tmpl, docs] = await Promise.all([
      listeningApi.sessions(),
      listeningApi.templates(),
      listeningApi.list(),
    ])
    sessions.value = sess
    templates.value = tmpl
    list.value = docs
  } finally {
    loading.value = false
  }
}

async function handleGenerate() {
  if (!form.session_id) {
    ElMessage.warning('请选择课次')
    return
  }
  if (!form.template_path) {
    ElMessage.warning('请选择听课模板')
    return
  }
  generating.value = true
  try {
    await listeningApi.generate({ session_id: form.session_id, template_path: form.template_path })
    ElMessage.success('生成成功')
    await loadAll()
  } finally {
    generating.value = false
  }
}

async function handleOpen(row) {
  await utilApi.openLocation({ offering_id: row.offering_id, kind: 'listening_record', document_id: row.id })
}

async function handleDelete(row) {
  try {
    await ElMessageBox.confirm(`确定删除「${row.course_name || ''}」的听课记录文档吗？`, '删除确认', {
      type: 'warning',
      confirmButtonText: '确定删除',
      cancelButtonText: '取消',
    })
  } catch {
    return
  }
  try {
    await listeningApi.delete(row.id)
    ElMessage.success('删除成功')
    await loadAll()
  } catch (e) {
    ElMessage.error('删除失败')
  }
}

onMounted(() => {
  loadAll()
})
</script>

<style scoped>
.mb16 {
  margin-bottom: 16px;
}
</style>

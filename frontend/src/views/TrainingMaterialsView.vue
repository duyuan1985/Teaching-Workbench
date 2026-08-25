<template>
  <div class="page-container">
    <!-- 生成面板 -->
    <div class="panel">
      <h2><el-icon><DocumentAdd /></el-icon> 生成实训资料</h2>
      <el-form :model="form" label-width="100px" v-loading="generating" style="max-width: 600px">
        <el-form-item label="实训课程">
          <el-select v-model="form.offering_id" placeholder="选择实训课程" filterable style="width: 100%">
            <el-option
              v-for="o in offerings"
              :key="o.id"
              :label="`${o.term} · ${o.course_name} · ${o.class_name || ''}`"
              :value="o.id"
            />
          </el-select>
        </el-form-item>
        <el-form-item label="资料目录">
          <el-input v-model="form.source_dir" placeholder="请输入资料目录路径，如 D:\教学资料\实训" clearable />
        </el-form-item>
        <el-form-item label="班级">
          <el-input v-model="form.class_name" placeholder="请输入班级名称" clearable />
        </el-form-item>
        <el-form-item>
          <el-button type="primary" :loading="generating" :icon="DocumentAdd" @click="handleGenerate">生成</el-button>
        </el-form-item>
      </el-form>
    </div>

    <!-- 文档列表 -->
    <div class="panel">
      <div class="actions mb16">
        <h2 style="margin-bottom: 0"><el-icon><FolderOpened /></el-icon> 实训资料文档</h2>
        <el-button :loading="loading" :icon="Refresh" @click="loadList">刷新</el-button>
      </div>
      <el-table :data="list" v-loading="loading" border stripe empty-text="暂无文档">
        <el-table-column prop="term" label="学期" width="120" />
        <el-table-column prop="course_name" label="课程" min-width="140" />
        <el-table-column prop="class_name" label="班级" width="120" />
        <el-table-column prop="source_dir" label="资料目录" min-width="200" show-overflow-tooltip />
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
import { Refresh, Delete, FolderOpened, DocumentAdd } from '@element-plus/icons-vue'
import { trainingApi, utilApi } from '../api'

const loading = ref(false)
const generating = ref(false)
const offerings = ref([])
const list = ref([])

const form = reactive({
  offering_id: null,
  source_dir: '',
  class_name: '',
})

async function loadOfferings() {
  offerings.value = await trainingApi.offerings()
}

async function loadList() {
  loading.value = true
  try {
    list.value = await trainingApi.list()
  } finally {
    loading.value = false
  }
}

async function handleGenerate() {
  if (!form.offering_id) {
    ElMessage.warning('请选择实训课程')
    return
  }
  if (!form.source_dir.trim()) {
    ElMessage.warning('请输入资料目录')
    return
  }
  generating.value = true
  try {
    await trainingApi.generate({
      offering_id: form.offering_id,
      source_dir: form.source_dir,
      class_name: form.class_name,
    })
    ElMessage.success('生成成功')
    form.source_dir = ''
    form.class_name = ''
    await loadList()
  } finally {
    generating.value = false
  }
}

async function handleOpen(row) {
  await utilApi.openLocation({ offering_id: row.offering_id, kind: 'training_material', document_id: row.id })
}

async function handleDelete(row) {
  try {
    await ElMessageBox.confirm(`确定删除「${row.course_name}」的实训资料文档吗？`, '删除确认', {
      type: 'warning',
      confirmButtonText: '确定删除',
      cancelButtonText: '取消',
    })
  } catch {
    return
  }
  try {
    await trainingApi.delete(row.id)
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

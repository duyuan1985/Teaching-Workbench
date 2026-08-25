<template>
  <div class="page-container">
    <!-- 新增/编辑表单 -->
    <div class="panel">
      <h2>
        <el-icon><EditPen /></el-icon>
        {{ editingId ? '编辑课程类型' : '新增课程类型' }}
      </h2>
      <el-form :model="form" label-width="80px" inline>
        <el-form-item label="名称">
          <el-input v-model="form.name" placeholder="请输入课程类型名称" style="width: 240px" clearable />
        </el-form-item>
        <el-form-item label="排序">
          <el-input-number v-model="form.sort_order" :min="0" :max="9999" controls-position="right" />
        </el-form-item>
        <el-form-item>
          <el-button type="primary" :loading="saving" @click="handleSave">
            {{ editingId ? '更新' : '新增' }}
          </el-button>
          <el-button v-if="editingId" @click="resetForm">取消编辑</el-button>
        </el-form-item>
      </el-form>
    </div>

    <!-- 列表 -->
    <div class="panel">
      <div class="actions mb16">
        <h2 style="margin-bottom: 0"><el-icon><List /></el-icon> 课程类型列表</h2>
        <el-button @click="loadList" :loading="loading" :icon="Refresh">刷新</el-button>
      </div>
      <el-table :data="list" v-loading="loading" border stripe empty-text="暂无课程类型">
        <el-table-column prop="name" label="名称" min-width="200" />
        <el-table-column prop="sort_order" label="排序" width="100" align="center" sortable />
        <el-table-column label="操作" width="160" align="center" fixed="right">
          <template #default="{ row }">
            <el-button size="small" :icon="Edit" @click="handleEdit(row)">编辑</el-button>
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
import { Edit, Delete, Refresh, List, EditPen } from '@element-plus/icons-vue'
import { courseTypesApi } from '../api'

const loading = ref(false)
const saving = ref(false)
const list = ref([])
const editingId = ref(null)

const form = reactive({
  name: '',
  sort_order: 0,
})

async function loadList() {
  loading.value = true
  try {
    list.value = await courseTypesApi.list()
  } finally {
    loading.value = false
  }
}

function resetForm() {
  editingId.value = null
  form.name = ''
  form.sort_order = 0
}

function handleEdit(row) {
  editingId.value = row.id
  form.name = row.name
  form.sort_order = row.sort_order ?? 0
}

async function handleSave() {
  if (!form.name.trim()) {
    ElMessage.warning('请输入课程类型名称')
    return
  }
  saving.value = true
  try {
    if (editingId.value) {
      await courseTypesApi.update(editingId.value, { name: form.name, sort_order: form.sort_order })
      ElMessage.success('更新成功')
    } else {
      await courseTypesApi.create({ name: form.name, sort_order: form.sort_order })
      ElMessage.success('新增成功')
    }
    resetForm()
    await loadList()
  } finally {
    saving.value = false
  }
}

async function handleDelete(row) {
  try {
    await ElMessageBox.confirm(`确定删除课程类型「${row.name}」吗？`, '删除确认', {
      type: 'warning',
      confirmButtonText: '确定删除',
      cancelButtonText: '取消',
    })
  } catch {
    return
  }
  try {
    await courseTypesApi.delete(row.id)
    ElMessage.success('删除成功')
    await loadList()
  } catch (e) {
    ElMessage.error('删除失败')
  }
}

onMounted(() => {
  loadList()
})
</script>

<style scoped>
.mb16 {
  margin-bottom: 16px;
}
</style>

<template>
  <div class="page-container">
    <!-- 筛选面板 -->
    <div class="panel">
      <h2><el-icon><Search /></el-icon> 筛选条件</h2>
      <el-form :model="filters" label-width="80px" inline>
        <el-form-item label="课程">
          <el-select v-model="filters.course" placeholder="全部课程" filterable clearable style="width: 280px">
            <el-option
              v-for="o in offerings"
              :key="o.id"
              :label="`${o.term || ''} · ${o.course_name || ''}`"
              :value="o.course_name || o.id"
            />
          </el-select>
        </el-form-item>
        <el-form-item label="资源类型">
          <el-select v-model="filters.kind" placeholder="全部类型" clearable style="width: 200px">
            <el-option
              v-for="t in resourceTypes"
              :key="t.value ?? t"
              :label="t.label ?? t"
              :value="t.value ?? t"
            />
          </el-select>
        </el-form-item>
        <el-form-item>
          <el-button type="primary" :icon="Search" @click="loadList">查询</el-button>
          <el-button :icon="RefreshLeft" @click="resetFilters">重置</el-button>
        </el-form-item>
      </el-form>
    </div>

    <!-- 资源列表 -->
    <div class="panel">
      <div class="actions mb16">
        <h2 style="margin-bottom: 0"><el-icon><Files /></el-icon> 教学资源列表</h2>
        <el-button :loading="loading" :icon="Refresh" @click="loadList">刷新</el-button>
      </div>
      <el-table :data="list" v-loading="loading" border stripe empty-text="暂无资源">
        <el-table-column prop="term" label="学期" width="120" />
        <el-table-column prop="course_name" label="课程" min-width="140" />
        <el-table-column prop="kind" label="类型" width="120">
          <template #default="{ row }">
            <el-tag size="small" type="info">{{ row.kind }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="name" label="名称" min-width="200" show-overflow-tooltip />
        <el-table-column prop="summary" label="内容摘要" min-width="300" show-overflow-tooltip />
      </el-table>
    </div>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { Search, Refresh, RefreshLeft, Files } from '@element-plus/icons-vue'
import { resourcesApi, offeringsApi } from '../api'

const loading = ref(false)
const offerings = ref([])
const resourceTypes = ref([])
const list = ref([])

const filters = reactive({
  course: '',
  kind: '',
})

async function loadFilters() {
  const [off, types] = await Promise.all([
    offeringsApi.list(),
    resourcesApi.types(),
  ])
  offerings.value = off
  resourceTypes.value = types
}

async function loadList() {
  loading.value = true
  try {
    list.value = await resourcesApi.list(filters.course || undefined, filters.kind || undefined)
  } finally {
    loading.value = false
  }
}

function resetFilters() {
  filters.course = ''
  filters.kind = ''
  loadList()
}

onMounted(() => {
  loadFilters()
  loadList()
})
</script>

<style scoped>
.mb16 {
  margin-bottom: 16px;
}
</style>

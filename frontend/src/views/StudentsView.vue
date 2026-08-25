<template>
  <div class="page-container">
    <!-- 班级选择 -->
    <div class="panel">
      <h2>班级学生名单</h2>
      <div class="filter-bar">
        <el-select
          v-model="selectedClass"
          placeholder="选择班级"
          clearable
          filterable
          class="filter-item"
          @change="onClassChange"
        >
          <el-option
            v-for="c in classes"
            :key="c"
            :label="c"
            :value="c"
          />
        </el-select>
        <el-button
          type="danger"
          plain
          :disabled="selectedRows.length === 0"
          @click="batchDelete"
        >
          批量删除 ({{ selectedRows.length }})
        </el-button>
      </div>
    </div>

    <!-- 学生表格 -->
    <div class="panel" v-loading="loading">
      <h2>
        学生列表
        <span v-if="selectedClass" class="badge">{{ selectedClass }}</span>
        <span class="muted" style="font-size: 13px">共 {{ students.length }} 人</span>
      </h2>
      <el-table
        :data="students"
        border
        stripe
        style="width: 100%"
        empty-text="暂无学生数据"
        @selection-change="onSelectionChange"
      >
        <el-table-column type="selection" width="45" align="center" />
        <el-table-column label="序号" type="index" width="60" align="center" />
        <el-table-column prop="student_no" label="学号" width="140" />
        <el-table-column label="姓名" width="160">
          <template #default="{ row }">
            <el-input
              v-if="editingRow === row.id"
              v-model="editForm.student_name"
              size="small"
              placeholder="姓名"
            />
            <span v-else>{{ row.student_name }}</span>
          </template>
        </el-table-column>
        <el-table-column label="性别" width="120">
          <template #default="{ row }">
            <el-input
              v-if="editingRow === row.id"
              v-model="editForm.gender"
              size="small"
              placeholder="性别"
            />
            <span v-else>{{ row.gender }}</span>
          </template>
        </el-table-column>
        <el-table-column prop="notes" label="备注" min-width="150" show-overflow-tooltip />
        <el-table-column label="操作" width="200" align="center">
          <template #default="{ row }">
            <template v-if="editingRow === row.id">
              <el-button
                type="primary"
                link
                size="small"
                :loading="row._saving"
                @click="saveEdit(row)"
              >
                保存
              </el-button>
              <el-button
                link
                size="small"
                @click="cancelEdit"
              >
                取消
              </el-button>
            </template>
            <template v-else>
              <el-button
                type="primary"
                link
                size="small"
                @click="startEdit(row)"
              >
                编辑
              </el-button>
              <el-button
                type="danger"
                link
                size="small"
                @click="deleteOne(row)"
              >
                删除
              </el-button>
            </template>
          </template>
        </el-table-column>
      </el-table>
    </div>

    <!-- 新增学生表单 -->
    <div class="panel">
      <h2>新增学生名单</h2>
      <p class="muted" style="margin-bottom: 12px">
        每行一名学生，格式：<code>学号 姓名 性别</code>（性别可省略）。例如：<code>2024001 张三 男</code>
      </p>
      <el-form :model="addForm" label-width="80px">
        <el-form-item label="班级">
          <el-input
            v-model="addForm.class_name"
            placeholder="输入班级名称，如：计科2301"
            style="width: 300px"
          />
        </el-form-item>
        <el-form-item label="学生名单">
          <el-input
            v-model="addForm.lines"
            type="textarea"
            :rows="8"
            placeholder="2024001 张三 男&#10;2024002 李四 女&#10;2024003 王五 男"
            style="width: 400px"
          />
        </el-form-item>
        <el-form-item>
          <el-button
            type="primary"
            :loading="adding"
            @click="addStudents"
          >
            保存名单
          </el-button>
        </el-form-item>
      </el-form>
    </div>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { studentsApi } from '../api'

// ---- 状态 ----
const classes = ref([])
const selectedClass = ref('')
const students = ref([])
const loading = ref(false)
const adding = ref(false)
const selectedRows = ref([])
const editingRow = ref(null)
const editForm = reactive({ student_name: '', gender: '' })

// ---- 新增表单 ----
const addForm = reactive({
  class_name: '',
  lines: '',
})

// ---- 加载学生列表 ----
async function loadStudents() {
  loading.value = true
  try {
    const res = await studentsApi.list(selectedClass.value)
    classes.value = res.classes || []
    if (res.selected_class) selectedClass.value = res.selected_class
    students.value = (res.students || []).map((s) => ({ ...s, _saving: false }))
  } catch (e) {
    // 已提示
  } finally {
    loading.value = false
  }
}

// ---- 班级切换 ----
async function onClassChange() {
  await loadStudents()
}

// ---- 表格选择 ----
function onSelectionChange(rows) {
  selectedRows.value = rows
}

// ---- 编辑单条（创建编辑副本，保存时才写回原数据）----
function startEdit(row) {
  editingRow.value = row.id
  editForm.student_name = row.student_name
  editForm.gender = row.gender
}

function cancelEdit() {
  editingRow.value = null
}

// ---- 保存编辑 ----
async function saveEdit(row) {
  if (!editForm.student_name) {
    ElMessage.warning('请输入姓名')
    return
  }
  row._saving = true
  try {
    await studentsApi.update(row.id, {
      student_name: editForm.student_name,
      gender: editForm.gender || '',
    })
    // 成功后写回原数据
    row.student_name = editForm.student_name
    row.gender = editForm.gender || ''
    ElMessage.success(`${editForm.student_name} 已保存`)
    editingRow.value = null
  } catch (e) {
    // 已提示
  } finally {
    row._saving = false
  }
}

// ---- 删除单条 ----
async function deleteOne(row) {
  try {
    await ElMessageBox.confirm(`确定删除学生"${row.student_name}"（${row.student_no}）？`, '提示', { type: 'warning' })
  } catch {
    return
  }
  try {
    await studentsApi.delete(row.id)
    ElMessage.success('已删除')
    await loadStudents()
  } catch (e) {
    ElMessage.error('删除失败')
  }
}

// ---- 批量删除 ----
async function batchDelete() {
  if (selectedRows.value.length === 0) return
  try {
    await ElMessageBox.confirm(
      `确定删除选中的 ${selectedRows.value.length} 名学生？`,
      '批量删除',
      { type: 'warning' },
    )
  } catch {
    return
  }
  try {
    for (const row of selectedRows.value) {
      await studentsApi.delete(row.id)
    }
    ElMessage.success(`已删除 ${selectedRows.value.length} 名学生`)
    selectedRows.value = []
    await loadStudents()
  } catch (e) {
    ElMessage.error('删除失败')
  }
}

// ---- 新增学生名单 ----
async function addStudents() {
  if (!addForm.class_name) {
    ElMessage.warning('请输入班级名称')
    return
  }
  if (!addForm.lines.trim()) {
    ElMessage.warning('请输入学生名单')
    return
  }
  adding.value = true
  try {
    const res = await studentsApi.create({
      class_name: addForm.class_name,
      lines: addForm.lines,
    })
    ElMessage.success(`已添加 ${res.added || 0} 名学生`)
    addForm.lines = ''
    // 自动切换到该班级并刷新
    selectedClass.value = addForm.class_name
    await loadStudents()
  } catch (e) {
    // 已提示
  } finally {
    adding.value = false
  }
}

// ---- 页面初始化 ----
onMounted(async () => {
  await loadStudents()
})
</script>

<style scoped>
.filter-bar {
  display: flex;
  gap: 12px;
  align-items: center;
  flex-wrap: wrap;
}

.filter-item {
  width: 280px;
}

code {
  background: #e3eaf1;
  color: var(--text);
  border-radius: 4px;
  padding: 1px 6px;
  font-size: 12px;
}
</style>

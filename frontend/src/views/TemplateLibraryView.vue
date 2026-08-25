<template>
  <div class="page-container">
    <!-- 模板库列表 -->
    <div class="panel">
      <div class="actions mb16">
        <h2 style="margin-bottom: 0"><el-icon><Files /></el-icon> 模板库</h2>
        <div>
          <el-button type="primary" :icon="FolderAdd" :loading="scanning" @click="handleScan">扫描模板目录</el-button>
          <el-button @click="loadList" :loading="loading" :icon="Refresh">刷新</el-button>
        </div>
      </div>
      <el-alert v-if="scanResult" :title="`扫描完成：新增 ${scanResult.added}，更新 ${scanResult.updated}，未变化 ${scanResult.skipped}`"
        type="success" show-icon closable class="mb16" @close="scanResult = null" />
      <el-table :data="list" v-loading="loading" border stripe empty-text="模板库为空，点击「扫描模板目录」导入">
        <el-table-column prop="doc_type" label="文档类型" width="110">
          <template #default="{ row }">
            <el-tag :type="typeTag(row.doc_type)" size="small">{{ row.doc_type }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column prop="name" label="模板名称" min-width="240" show-overflow-tooltip />
        <el-table-column prop="version_label" label="版本" width="110">
          <template #default="{ row }">{{ row.version_label || '—' }}</template>
        </el-table-column>
        <el-table-column prop="file_format" label="格式" width="70" align="center" />
        <el-table-column label="文件" width="70" align="center">
          <template #default="{ row }">
            <el-tag :type="row.exists ? 'success' : 'danger'" size="small">{{ row.exists ? '在' : '失' }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column label="解析状态" width="100" align="center">
          <template #default="{ row }">
            <el-tag :type="statusTag(row.status)" size="small">{{ row.status }}</el-tag>
          </template>
        </el-table-column>
        <el-table-column label="契约" width="170">
          <template #default="{ row }">
            <template v-if="row.contract_id">
              <el-tag :type="row.contract_status === '已确认' ? 'success' : 'warning'" size="small">
                {{ row.contract_status }} v{{ row.contract_version }}
              </el-tag>
              <span class="slot-info">槽位 {{ row.slot_count }}</span>
            </template>
            <span v-else class="slot-info">未解析</span>
          </template>
        </el-table-column>
        <el-table-column label="操作" width="250" align="center" fixed="right">
          <template #default="{ row }">
            <el-button size="small" type="primary" :icon="MagicStick" :loading="parsingId === row.id"
              :disabled="row.status === '暂不支持'" @click="handleParse(row)">解析</el-button>
            <el-button size="small" :icon="View" :disabled="!row.contract_id" @click="openWorkbench(row)">契约</el-button>
            <el-button size="small" type="danger" :icon="Delete" @click="handleDelete(row)">删除</el-button>
          </template>
        </el-table-column>
      </el-table>
    </div>

    <!-- 契约确认工作台 -->
    <div class="panel" v-if="activeTemplate">
      <div class="actions mb16">
        <h2 style="margin-bottom: 0">
          <el-icon><DocumentChecked /></el-icon>
          契约工作台：{{ activeTemplate.name }}
        </h2>
        <div>
          <el-button @click="activeTemplate = null" :icon="Close">关闭</el-button>
          <el-button type="success" :icon="Check" :loading="confirming" :disabled="!contract || contract.status === '已确认'"
            @click="handleConfirm">确认契约</el-button>
        </div>
      </div>

      <div v-if="contractLoading" v-loading="true" style="height: 200px"></div>
      <template v-else-if="contract">
        <el-alert v-if="contract.status === '已确认'" type="success" show-icon class="mb16" :closable="false"
          :title="`该契约已于 ${contract.confirmed_at} 确认（v${contract.version}），重新解析将生成新版本`" />
        <el-alert v-else type="info" show-icon class="mb16" :closable="false"
          title="核对下方三契约：结构（表格角色/循环）、格式（字体/下划线）、内容（指令/参考格式/强制条款）。槽位可逐条修正后确认。" />

        <el-tabs v-model="tab">
          <el-tab-pane label="槽位契约" name="slots">
            <div class="mb16" style="color: var(--el-text-color-secondary); font-size: 13px">
              共 {{ contract.slots.length }} 个槽位：A=事实提取，B=AI润色，C=结构生成，人工=手工填写。
              低置信度槽位须逐一核对后方可确认契约。
            </div>
            <el-table :data="contract.slots" border stripe size="small" max-height="480">
              <el-table-column prop="section_title" label="章节" min-width="120" show-overflow-tooltip />
              <el-table-column label="字段" min-width="120">
                <template #default="{ row }">
                  <span :class="{ 'manual-mark': row.manual_override }">{{ row.field_name }}</span>
                </template>
              </el-table-column>
              <el-table-column label="分类" width="130">
                <template #default="{ row }">
                  <el-select :model-value="row.classification" size="small" @change="(v) => patchSlot(row, 'classification', v)">
                    <el-option label="A 事实" value="A" />
                    <el-option label="B 润色" value="B" />
                    <el-option label="C 生成" value="C" />
                    <el-option label="人工" value="人工" />
                  </el-select>
                </template>
              </el-table-column>
              <el-table-column label="置信度" width="90">
                <template #default="{ row }">
                  <el-tag :type="confTag(row.confidence)" size="small" effect="plain">{{ row.confidence }}</el-tag>
                </template>
              </el-table-column>
              <el-table-column label="格式契约" min-width="200">
                <template #default="{ row }">
                  <div class="fmt">{{ formatSummary(row.format_json) }}</div>
                </template>
              </el-table-column>
              <el-table-column prop="content_req" label="内容要求" min-width="200" show-overflow-tooltip>
                <template #default="{ row }">
                  <span v-if="row.content_req">{{ row.content_req }}</span>
                  <span v-else style="color: var(--el-text-color-placeholder)">—</span>
                </template>
              </el-table-column>
              <el-table-column prop="structure_json.repeat_scope" label="循环" min-width="140" show-overflow-tooltip />
              <el-table-column label="核对" width="90" align="center">
                <template #default="{ row }">
                  <el-button size="small" :type="row.approval_status === '已确认' ? 'success' : 'warning'" link
                    @click="patchSlot(row, 'approval_status', row.approval_status === '已确认' ? '待确认' : '已确认')">
                    {{ row.approval_status === '已确认' ? '已核对' : '核对' }}
                  </el-button>
                </template>
              </el-table-column>
            </el-table>
          </el-tab-pane>

          <el-tab-pane label="内容契约" name="content">
            <h3 class="section-title">格式指令（模板明确规定的字体/行距要求）</h3>
            <el-empty v-if="!contract.content_json.format_instructions?.length" description="无" :image-size="48" />
            <el-table v-else :data="contract.content_json.format_instructions" border size="small">
              <el-table-column prop="scope" label="作用范围" width="130">
                <template #default="{ row }">
                  <el-tag size="small" :type="row.scope === '全文档正文' ? 'danger' : 'info'">{{ row.scope }}</el-tag>
                </template>
              </el-table-column>
              <el-table-column prop="section" label="所在章节" width="200" show-overflow-tooltip />
              <el-table-column prop="text" label="指令内容" min-width="320" />
            </el-table>
            <h3 class="section-title">强制条款（来自官方完善要求）</h3>
            <el-empty v-if="!contract.content_json.mandates?.length" description="无" :image-size="48" />
            <ul v-else class="plain-list">
              <li v-for="(m, i) in contract.content_json.mandates" :key="i">{{ m }}</li>
            </ul>
            <h3 class="section-title">参考格式</h3>
            <el-empty v-if="!contract.content_json.reference_formats?.length" description="未识别到参考格式" :image-size="48" />
            <el-table v-else :data="contract.content_json.reference_formats" border size="small">
              <el-table-column prop="locator" label="位置" min-width="160" />
              <el-table-column label="格式类型与要点" min-width="380">
                <template #default="{ row }">
                  <div v-for="f in row.formats" :key="f.type" class="ref-fmt">
                    <el-tag size="small" type="info">{{ refTypeName(f.type) }}</el-tag>
                    <span>{{ refDetail(f) }}</span>
                  </div>
                </template>
              </el-table-column>
            </el-table>
            <h3 class="section-title">模板指令</h3>
            <el-empty v-if="!contract.content_json.instructions?.length" description="无" :image-size="48" />
            <ul v-else class="plain-list">
              <li v-for="(item, i) in contract.content_json.instructions" :key="i">
                <el-tag size="small" effect="plain" class="loc-tag">{{ item.locator }}</el-tag> {{ item.text }}
              </li>
            </ul>
          </el-tab-pane>

          <el-tab-pane label="结构契约" name="structure">
            <h3 class="section-title">封面字段（{{ contract.structural_json.cover?.fields?.length || 0 }} 项，提议：值加下划线并居中）</h3>
            <div class="cover-fields">
              <el-tag v-for="f in contract.structural_json.cover?.fields || []" :key="f.locator" class="cover-tag" size="small">
                {{ f.label }} <span class="cover-loc">{{ f.locator }}</span>
              </el-tag>
            </div>
            <h3 class="section-title">课程名称标题占位（{{ contract.structural_json.cover?.title_fields?.length || 0 }} 处）</h3>
            <el-table :data="contract.structural_json.cover?.title_fields || []" border size="small">
              <el-table-column label="归属" width="70">
                <template #default="{ row }">
                  <el-tag :type="row.page === '封面' ? 'primary' : 'success'" size="small">{{ row.page }}</el-tag>
                </template>
              </el-table-column>
              <el-table-column label="所在章节" prop="section" width="170" show-overflow-tooltip>
                <template #default="{ row }">{{ row.section || '—' }}</template>
              </el-table-column>
              <el-table-column prop="placeholder" label="占位符" width="130" />
              <el-table-column prop="fill_template" label="填充后" min-width="200">
                <template #default="{ row }">
                  <code>{{ row.fill_template }}</code>
                </template>
              </el-table-column>
              <el-table-column label="需删除的模板提示" min-width="260">
                <template #default="{ row }">
                  <span v-if="row.clean_instructions.length">{{ row.clean_instructions.join('、') }}</span>
                  <span v-else style="color: var(--el-text-color-placeholder)">无</span>
                </template>
              </el-table-column>
            </el-table>
            <h3 class="section-title">表格结构（{{ contract.structural_json.tables.length }} 个）</h3>
            <el-table :data="contract.structural_json.tables" border size="small">
              <el-table-column prop="index" label="#" width="50" align="center" />
              <el-table-column prop="role" label="角色" min-width="140" />
              <el-table-column label="尺寸" width="90" align="center">
                <template #default="{ row }">{{ row.rows }}×{{ row.columns }}</template>
              </el-table-column>
              <el-table-column label="合并单元格" width="90" align="center">
                <template #default="{ row }">{{ row.merged_cells.length }} 处</template>
              </el-table-column>
              <el-table-column prop="repeat_mode" label="循环模式" min-width="180" />
              <el-table-column prop="header_text" label="表头摘要" min-width="240" show-overflow-tooltip />
            </el-table>
          </el-tab-pane>
        </el-tabs>
      </template>
    </div>
  </div>
</template>

<script setup>
import { ref, onMounted } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import { Delete, Refresh, View, MagicStick, FolderAdd, Files, DocumentChecked, Close, Check } from '@element-plus/icons-vue'
import { templateLibraryApi } from '../api'

const loading = ref(false)
const scanning = ref(false)
const parsingId = ref(null)
const confirming = ref(false)
const list = ref([])
const scanResult = ref(null)
const activeTemplate = ref(null)
const contract = ref(null)
const contractLoading = ref(false)
const tab = ref('slots')

function typeTag(t) {
  return { 课程标准: 'primary', 教学设计: 'success', 授课计划: 'warning', 完善要求: 'info', 听课记录: 'info', 其他: 'info' }[t] || 'info'
}
function statusTag(s) {
  return { 已解析: 'success', 解析失败: 'danger', 暂不支持: 'info' }[s] || 'warning'
}
function confTag(c) {
  return { 高: 'success', 中: 'warning', 低: 'danger' }[c] || 'info'
}
function refTypeName(t) {
  return { aspect_list: '方面列举', numbered_placeholder: '编号要点', definition_pattern: '定义句式' }[t] || t
}
function refDetail(f) {
  if (f.aspects) return f.aspects.join('、') + ' 等方面'
  if (f.items) return f.items.slice(0, 3).join('；')
  return f.example || ''
}
function formatSummary(fmt) {
  if (!fmt || Object.keys(fmt).length === 0) return '（继承模板）'
  const parts = []
  if (fmt.underline) parts.push('下划线')
  if (fmt.font_name || fmt.east_asia_font) parts.push(fmt.east_asia_font || fmt.font_name)
  if (fmt.font_size_pt) parts.push(`${fmt.font_size_pt}pt`)
  if (fmt.bold) parts.push('加粗')
  if (fmt.alignment === 'center' || fmt.alignment === 1) parts.push('居中')
  return parts.join(' · ') || '（继承模板）'
}

async function loadList() {
  loading.value = true
  try {
    list.value = await templateLibraryApi.list()
  } finally {
    loading.value = false
  }
}

async function handleScan() {
  scanning.value = true
  try {
    scanResult.value = await templateLibraryApi.scan()
    ElMessage.success('扫描完成')
    await loadList()
  } finally {
    scanning.value = false
  }
}

async function handleParse(row) {
  parsingId.value = row.id
  try {
    const out = await templateLibraryApi.parse(row.id)
    ElMessage.success(out.status === '已解析' ? `解析成功：${out.message}` : out.message)
    await loadList()
    if (activeTemplate.value?.id === row.id) openWorkbench({ ...row })
    else if (out.status === '已解析') openWorkbench({ ...row, contract_id: out.contract_id })
  } catch (e) {
    ElMessage.error(typeof e === 'string' ? e : '解析失败')
  } finally {
    parsingId.value = null
  }
}

async function openWorkbench(row) {
  activeTemplate.value = row
  contractLoading.value = true
  contract.value = null
  tab.value = 'slots'
  try {
    contract.value = await templateLibraryApi.contract(row.id)
  } catch (e) {
    ElMessage.error('契约加载失败')
    activeTemplate.value = null
  } finally {
    contractLoading.value = false
  }
}

async function patchSlot(row, field, value) {
  try {
    await templateLibraryApi.updateSlot(row.id, { [field]: value })
    row[field] = value
    row.manual_override = 1
    if (field === 'approval_status' && value === '已确认') row.confidence = '高'
  } catch (e) {
    ElMessage.error('保存失败')
  }
}

async function handleConfirm() {
  const pending = contract.value.slots.filter((s) => s.confidence === '低' && s.approval_status !== '已确认')
  if (pending.length) {
    ElMessage.warning(`还有 ${pending.length} 个低置信度槽位未核对`)
    return
  }
  try {
    await ElMessageBox.confirm('确认后该契约将作为模板生成的依据，重新解析会生成新版本。', '确认契约', { type: 'warning' })
  } catch {
    return
  }
  confirming.value = true
  try {
    await templateLibraryApi.confirmContract(contract.value.id)
    ElMessage.success('契约已确认')
    await loadList()
    await openWorkbench(activeTemplate.value)
  } finally {
    confirming.value = false
  }
}

async function handleDelete(row) {
  try {
    await ElMessageBox.confirm(`确定从模板库删除「${row.name}」吗？（不删除磁盘文件）`, '删除确认', {
      type: 'warning', confirmButtonText: '确定删除', cancelButtonText: '取消',
    })
  } catch {
    return
  }
  try {
    await templateLibraryApi.delete(row.id)
    ElMessage.success('已删除')
    if (activeTemplate.value?.id === row.id) activeTemplate.value = null
    await loadList()
  } catch (e) {
    ElMessage.error(typeof e === 'string' ? e : '删除失败')
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
.actions {
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.slot-info {
  margin-left: 6px;
  font-size: 12px;
  color: var(--el-text-color-secondary);
}
.section-title {
  margin: 18px 0 10px;
  font-size: 14px;
  color: var(--el-text-color-primary);
}
.plain-list {
  margin: 0;
  padding-left: 18px;
  line-height: 1.9;
  font-size: 13px;
}
.loc-tag {
  font-family: monospace;
  margin-right: 4px;
}
.fmt {
  font-size: 12px;
  color: var(--el-text-color-secondary);
}
.manual-mark {
  color: var(--el-color-warning);
}
.cover-fields {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}
.cover-tag .cover-loc {
  margin-left: 4px;
  font-family: monospace;
  font-size: 11px;
  opacity: 0.6;
}
.ref-fmt {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 4px;
}
</style>

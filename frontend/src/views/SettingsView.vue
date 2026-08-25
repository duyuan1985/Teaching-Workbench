<template>
  <div class="page-container">
    <el-row :gutter="20">
      <el-col :span="12">
        <!-- 本地AI面板 -->
        <div class="panel">
          <h2><el-icon><Cpu /></el-icon> 本地AI（Ollama）</h2>
          <el-skeleton v-if="loading" :rows="4" animated />
          <template v-else>
            <el-descriptions :column="2" border size="small" class="mb16">
              <el-descriptions-item label="Ollama 状态">
                <el-tag :type="settings.ollama_available ? 'success' : 'danger'" size="small">
                  {{ settings.ollama_available ? '在线' : '离线' }}
                </el-tag>
              </el-descriptions-item>
              <el-descriptions-item label="已安装模型数">
                <span class="badge">{{ settings.installed_models ?? 0 }}</span>
              </el-descriptions-item>
              <el-descriptions-item label="缓存数">
                <span class="badge">{{ settings.cache_count ?? 0 }}</span>
              </el-descriptions-item>
              <el-descriptions-item label="教师姓名">
                {{ settings.teacher_name || '-' }}
              </el-descriptions-item>
            </el-descriptions>

            <el-form label-width="120px" class="settings-form">
              <el-form-item label="Ollama 地址">
                <el-input v-model="form.ollama_url" placeholder="http://localhost:11434" clearable />
              </el-form-item>
              <el-form-item label="默认模型">
                <el-select v-model="form.ollama_model" placeholder="选择模型" style="width: 100%" clearable>
                  <el-option
                    v-for="m in installedModelList"
                    :key="m"
                    :label="m"
                    :value="m"
                  />
                </el-select>
                <div v-if="!installedModelList.length" class="muted">未检测到已安装模型，请先在 Ollama 中拉取模型</div>
              </el-form-item>
              <el-form-item label="教师姓名">
                <el-input v-model="form.teacher_name" placeholder="导入教学安排表时自动提取，也可手动修改" clearable />
                <div class="muted">导入教学安排表时自动从教师列提取，修改后重新生成文档即可生效</div>
              </el-form-item>
              <el-form-item label="输出根目录">
                <el-input v-model="form.output_root" placeholder="生成文档的保存路径" clearable />
              </el-form-item>
              <el-form-item>
                <el-button type="primary" :loading="saving" @click="handleSave">保存设置</el-button>
              </el-form-item>
            </el-form>
          </template>
        </div>
      </el-col>

      <el-col :span="12">
        <!-- AI增强生成面板 -->
        <div class="panel">
          <h2><el-icon><MagicStick /></el-icon> AI 增强生成</h2>
          <el-skeleton v-if="loading" :rows="3" animated />
          <template v-else>
            <el-descriptions :column="1" border size="small" class="mb16">
              <el-descriptions-item label="当前状态">
                <el-tag :type="settings.enhanced_generation ? 'success' : 'info'" size="small">
                  {{ settings.enhanced_generation ? '已启用' : '未启用' }}
                </el-tag>
              </el-descriptions-item>
              <el-descriptions-item label="已生成增强内容数">
                <span class="badge">{{ settings.enhanced_count ?? 0 }}</span>
              </el-descriptions-item>
            </el-descriptions>

            <el-button
              :type="settings.enhanced_generation ? 'warning' : 'success'"
              :loading="toggling"
              @click="handleToggleEnhanced"
            >
              {{ settings.enhanced_generation ? '关闭增强模式' : '启用增强模式' }}
            </el-button>
          </template>
        </div>

        <!-- 课程类型管理入口 -->
        <div class="panel">
          <h2><el-icon><FolderOpened /></el-icon> 课程类型管理</h2>
          <p class="muted mb16">管理课程分类，用于课程实例的归类与筛选。</p>
          <el-button type="primary" @click="$router.push('/course-types')">
            前往管理
            <el-icon class="el-icon--right"><ArrowRight /></el-icon>
          </el-button>
        </div>
      </el-col>
      <el-col :span="12">
        <div class="panel">
          <h2><el-icon><Setting /></el-icon> 蓝本审查规则</h2>
          <p class="muted mb16">
            为不同课程配置PPT分组方式、学习目标提取关键字等规则，或启用AI自动审查。
          </p>
          <el-button type="primary" @click="$router.push('/review-rules')">
            前往配置
            <el-icon class="el-icon--right"><ArrowRight /></el-icon>
          </el-button>
        </div>
      </el-col>
    </el-row>

    <!-- 使用原则说明 -->
    <div class="panel">
      <h2><el-icon><InfoFilled /></el-icon> 使用原则</h2>
      <ul class="principle-list">
        <li>所有教学档案数据均存储在本地，不会上传到云端。</li>
        <li>系统不会修改原始 Word 模板文件，所有生成内容均独立保存。</li>
        <li>AI 增强生成依赖本地 Ollama 服务，请确保 Ollama 已启动并安装所需模型。</li>
        <li>增强内容生成后会缓存，重复生成相同内容时可直接复用，减少等待时间。</li>
        <li>输出根目录：<span class="badge">{{ settings.output_root || '-' }}</span></li>
      </ul>
    </div>
  </div>
</template>

<script setup>
import { ref, reactive, onMounted, computed } from 'vue'
import { ElMessage } from 'element-plus'
import { settingsApi } from '../api'

const loading = ref(false)
const saving = ref(false)
const toggling = ref(false)

const settings = ref({
  ollama_url: '',
  ollama_model: '',
  ollama_available: false,
  installed_models: 0,
  cache_count: 0,
  enhanced_generation: false,
  enhanced_count: 0,
  teacher_name: '',
  output_root: '',
})

const form = reactive({
  ollama_url: '',
  ollama_model: '',
  teacher_name: '',
  output_root: '',
})

const installedModelList = computed(() => {
  const models = settings.value.installed_models
  if (Array.isArray(models)) return models
  if (typeof models === 'number') return []
  return []
})

async function loadSettings() {
  loading.value = true
  try {
    const data = await settingsApi.get()
    settings.value = { ...settings.value, ...data }
    form.ollama_url = data.ollama_url || ''
    form.ollama_model = data.ollama_model || ''
    form.teacher_name = data.teacher_name || ''
    form.output_root = data.output_root || ''
  } finally {
    loading.value = false
  }
}

async function handleSave() {
  saving.value = true
  try {
    await settingsApi.update({
      ollama_url: form.ollama_url,
      ollama_model: form.ollama_model,
      teacher_name: form.teacher_name,
      output_root: form.output_root,
    })
    ElMessage.success('设置已保存')
    await loadSettings()
  } finally {
    saving.value = false
  }
}

async function handleToggleEnhanced() {
  toggling.value = true
  try {
    await settingsApi.toggleEnhanced()
    ElMessage.success('增强模式已切换')
    await loadSettings()
  } finally {
    toggling.value = false
  }
}

onMounted(() => {
  loadSettings()
})
</script>

<style scoped>
.settings-form {
  max-width: 460px;
}

.mb16 {
  margin-bottom: 16px;
}

.principle-list {
  list-style: disc;
  padding-left: 20px;
  line-height: 2;
  color: var(--text);
}

.principle-list li {
  font-size: 13px;
}
</style>

import { ref } from 'vue'
import { ElMessage } from 'element-plus'

export function useApiList(fetchFn) {
  const data = ref([])
  const loading = ref(false)

  async function load() {
    loading.value = true
    try {
      data.value = await fetchFn()
    } catch (e) {
      ElMessage.error('加载数据失败')
    } finally {
      loading.value = false
    }
  }

  return { data, loading, load }
}

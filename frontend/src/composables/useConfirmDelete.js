import { ElMessage, ElMessageBox } from 'element-plus'

export function useConfirmDelete(deleteFn, { message = '确认删除？', successMsg = '已删除' } = {}) {
  async function confirmDelete(id, msg) {
    try {
      await ElMessageBox.confirm(msg || message, '删除确认', { type: 'warning' })
    } catch {
      return false
    }
    try {
      await deleteFn(id)
      ElMessage.success(successMsg)
      return true
    } catch (e) {
      ElMessage.error('删除失败')
      return false
    }
  }

  return { confirmDelete }
}

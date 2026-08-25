import { utilApi } from '../api'

export function useOpenFolder() {
  async function openLocation(offeringId, kind = '', documentId = 0) {
    try {
      await utilApi.openLocation({ offering_id: offeringId, kind, document_id: documentId })
    } catch (e) {
      // 静默失败
    }
  }
  return { openLocation }
}

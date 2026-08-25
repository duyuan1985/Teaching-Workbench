import { createRouter, createWebHistory } from 'vue-router'

const routes = [
  { path: '/', name: 'home', component: () => import('../views/HomeView.vue') },
  { path: '/offering/:id', name: 'offering', component: () => import('../views/OfferingView.vue') },
  { path: '/attendance', name: 'attendance', component: () => import('../views/AttendanceView.vue') },
  { path: '/students', name: 'students', component: () => import('../views/StudentsView.vue') },
  { path: '/grade-analysis', name: 'gradeAnalysis', component: () => import('../views/GradeAnalysisView.vue') },
  { path: '/training-materials', name: 'trainingMaterials', component: () => import('../views/TrainingMaterialsView.vue') },
  { path: '/listening-records', name: 'listeningRecords', component: () => import('../views/ListeningRecordsView.vue') },
  { path: '/settings', name: 'settings', component: () => import('../views/SettingsView.vue') },
  { path: '/template-library', name: 'templateLibrary', component: () => import('../views/TemplateLibraryView.vue') },
  { path: '/review-rules', name: 'reviewRules', component: () => import('../views/ReviewRulesView.vue') },
  { path: '/course-types', name: 'courseTypes', component: () => import('../views/CourseTypesView.vue') },
  { path: '/resources', name: 'resources', component: () => import('../views/ResourcesView.vue') },
  { path: '/assignment/:id', name: 'assignment', component: () => import('../views/AssignmentView.vue') },
  { path: '/:pathMatch(.*)*', name: 'notFound', component: () => import('../views/NotFoundView.vue') },
]

const router = createRouter({
  history: createWebHistory(),
  routes,
})

export default router

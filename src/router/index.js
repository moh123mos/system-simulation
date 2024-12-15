import { createRouter, createWebHistory } from 'vue-router'
import HomeView from '../views/HomeView.vue'
import CreateAnalysis from '@/views/createAnalysis.vue'
import ViewAnalysis from '@/views/viewAnalysis.vue'
import AiAnalyzer from '@/views/aiAnalyzer.vue'
import OurServices from '@/views/ourServices.vue'
const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      name: 'home',
      component: HomeView
    },
    {
      path: '/:level/create-analysis/',
      name: 'create-page',
      component: CreateAnalysis
    },
    {
      path: '/:level/view-analysis/',
      name: 'view-page',
      component: ViewAnalysis
    },
    {
      path: '/:level/ai-analyzer/',
      name: 'ai-page',
      component: AiAnalyzer
    },
    {
      path: '/:level/our-services/',
      name: 'services-page',
      component: OurServices
    },
  ]
})

export default router

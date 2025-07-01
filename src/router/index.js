import { createRouter, createWebHistory} from 'vue-router'
import HomeView from '../views/HomeView.vue'
import CreateAnalysis from '@/views/createAnalysis.vue'
import ViewAnalysis from '@/views/viewAnalysis.vue'
import AiAnalyzer from '@/views/aiAnalyzer.vue'
import OurServices from '@/views/ourServices.vue'
import AnimatorPage from '@/views/animatorPage.vue'

const routes = [
    {
      path: '/',
      name: 'home',
      component: HomeView,
    },
    {
      path: '/:level/our-services/',
      name: 'services-page',
      component: OurServices,
    },
    {
      path: '/:level/create-analysis/',
      name: 'create-page',
      component: CreateAnalysis,
    },
    {
      path: '/:level/view-analysis/',
      name: 'view-page',
      component: ViewAnalysis,
    },
    {
      path: '/:level/animator/',
      name: 'animation',
      component: AnimatorPage,
    },
    {
      path: '/:level/ai-analyzer/',
      name: 'ai-page',
      component: AiAnalyzer,
    },
  ]; 
  
const router = createRouter({
  history: createWebHistory(),
  routes,
})

export default router

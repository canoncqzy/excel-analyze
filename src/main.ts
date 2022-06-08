import { createApp } from 'vue'
import App from './App.vue'
import router from './router'
import store from './store'
import 'ant-design-vue/dist/antd.css'
import { Upload, message } from 'ant-design-vue'

const app = createApp(App)

app.use(Upload).use(store).use(router).mount('#app')
app.config.globalProperties.$message = message

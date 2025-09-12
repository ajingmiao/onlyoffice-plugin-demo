import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'

// https://vite.dev/config/
export default defineConfig({
  plugins: [vue()],
  server: {
    host: '0.0.0.0', // 监听所有网络接口，包括局域网IP
    port: 5174,      // 确保端口号正确
    cors: true       // 启用CORS支持
  }
})

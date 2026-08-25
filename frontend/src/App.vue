<template>
  <el-container class="app-container">
    <el-header class="app-header">
      <div class="header-inner">
        <div class="nav-links">
          <router-link to="/" class="nav-brand">教学档案工作台</router-link>
          <router-link to="/" class="nav-link">课程档案</router-link>
          <router-link to="/attendance" class="nav-link">上课</router-link>
          <router-link to="/students" class="nav-link">班级学生名单</router-link>
          <router-link to="/grade-analysis" class="nav-link">成绩分析</router-link>
          <router-link to="/training-materials" class="nav-link">实训资料</router-link>
          <router-link to="/listening-records" class="nav-link">听课记录</router-link>
          <router-link to="/template-library" class="nav-link">模板库</router-link>
          <router-link to="/settings" class="nav-link">基础设置</router-link>
        </div>
        <span class="header-tag">本地数据 · 不修改原Word</span>
      </div>
    </el-header>
    <el-main class="app-main">
      <router-view />
    </el-main>
  </el-container>
</template>

<script setup>
import { RouterView, RouterLink } from 'vue-router'
</script>

<style scoped>
.app-container {
  min-height: 100vh;
}

.app-header {
  background: var(--header-bg);
  padding: 0;
  height: 52px !important;
  position: sticky;
  top: 0;
  z-index: 100;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
}

.header-inner {
  max-width: 1400px;
  margin: 0 auto;
  padding: 0 20px;
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.nav-links {
  display: flex;
  align-items: center;
  gap: 4px;
}

.nav-brand {
  color: #fff;
  font-size: 16px;
  font-weight: bold;
  margin-right: 20px;
  text-decoration: none;
}

.nav-link {
  color: rgba(255, 255, 255, 0.7);
  padding: 6px 12px;
  border-radius: 6px;
  font-size: 14px;
  transition: all 0.2s;
}

.nav-link:hover {
  color: #fff;
  background: rgba(255, 255, 255, 0.1);
  text-decoration: none;
}

.router-link-active.nav-link {
  color: #fff;
  background: rgba(255, 255, 255, 0.15);
}

.header-tag {
  color: rgba(255, 255, 255, 0.5);
  font-size: 12px;
}

.app-main {
  padding: 0;
  max-width: 1400px;
  margin: 0 auto;
  width: 100%;
}
</style>

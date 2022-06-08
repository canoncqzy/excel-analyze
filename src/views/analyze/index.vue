<template>
  <div class="container">
    {{ title }}
  </div>
</template>

<script setup lang="ts">
import { ref } from 'vue'
const title = ref<string>('analyze1')
</script>

<style scoped lang="less"></style>

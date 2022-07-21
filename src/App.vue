<template>
  <div id="app">
    <luckysheet></luckysheet>
  </div>
</template>

<script>
import Luckysheet from './components/Luckysheet'

export default {
  name: 'App',
  components: {
    Luckysheet
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>

<template>
  <div id="app">
    <VistaPrincipal />
  </div>
</template>

<script>
import VistaPrincipal from './components/Vista.vue';

export default {
  name: 'App',
  components: {
    VistaPrincipal
  }
};
</script>

<style>
#app {
  font-family: 'Avenir', Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  height: 50%;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>

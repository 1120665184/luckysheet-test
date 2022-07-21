const { defineConfig } = require('@vue/cli-service')
module.exports = defineConfig({
  transpileDependencies: true,
  devServer:{
    proxy:{
      '/luckysheet-service':{
        target:'http://127.0.0.1:8888',
        pathRewrite:{'^/luckysheet-service':''},
        changeOrigin:true
      },
      '/socket-luckysheet-service':{
        target:'ws://127.0.0.1:8888',
        pathRewrite:{'^/socket-luckysheet-service':''},
        ws:true,
        changeOrigin:true
      }
    }
  }
})

<template>
  <div>
    <div class="workList">
      <h3>工作簿列表：</h3>
      <ul>
        <li v-for="(v,index) of listVo.records" :key="index"
            :class="{
              clicked : v.gridKey === gridKey
            }"
            @click="gridKeyChange(v.gridKey)">{{v.title}}</li>
      </ul>
    </div>
    <div class="workArea">
      <div class="toolbar">
        <button type="button" @click="edit" >插入或修改</button>
<!--        <button type="button" @click="getAllSheets">获取所有sheet</button>-->
        <button type="button" @click="deleteWorkbook">删除该工作簿</button>
      </div>
      <div id="luckysheet" style="height: 570px;width: 100%"></div>
    </div>
  </div>

</template>

<script>
import {saveOrEdit,getList,findDetail,deleteByGridKey} from '@/api/luckysheet'
const proxyPre = 'luckysheet-service'
const loadUrl = `/${proxyPre}/luckysheet/load`
const loadSheetUrl = `/${proxyPre}/luckysheet/loadSheet`
const updateUrl = `ws:${location.host}/socket-${proxyPre}/ws/lucksheet`
function initSheetOptions (){
  return {
    lang:'zh',
    title: '新建标题',
    column: 20,
    row: 50
  }
}
export default {
  name: "LuckysheetModel",
  data(){
    return {
      // eslint-disable-next-line no-undef
      sheet: luckysheet,
      gridKey:'',
      luckysheetOptions:initSheetOptions(),
      listForm:{
        pageNumber : 1,
        pageSize : 100 ,
        title:''
      },
      listVo:{
        records:[],
        total:0
      }
    }
  },
  computed:{
    luckysheet(){
      let options = this.luckysheetOptions
      if(this.gridKey){
        options.gridKey = this.gridKey
        options.loadUrl = loadUrl
        options.loadSheetUrl = loadSheetUrl
        options.allowUpdate = true
        options.updateUrl = updateUrl
      } else {
        options = initSheetOptions()
      }
      return options
    }
  },
  methods:{
    createSheet:function (dom){
      let options = {
        container:dom,
        ...this.luckysheet,
      }
      this.sheet.create(options)
    },
    edit: function (){
      saveOrEdit(this.sheet.toJson())
      .then(res =>{
        if(res.status === 200){
          alert("保存成功")
          this.getList()
          this.gridKey = res.data
        }else {
          alert("保存失败")
        }
      })
    },
    getAllSheets:function (){
      console.log()
    },
    getList:function(){
      getList(this.listForm).then(res =>{
        if(res.status === 200){
          this.listVo = res.data
        }else {
          alert("错误")
        }
      })
    },
    gridKeyChange:function (gridKey){
      console.log(gridKey)
      this.gridKey = gridKey;
    },
    getDetail:function(func){
      findDetail(this.gridKey).then(res => {
        if(200 === res.status){
          this.luckysheetOptions.title = res.data.title
          func()
        }
      })
    },
    deleteWorkbook:function (){
      if(!this.gridKey){
        alert("请选择一个工作簿!")
        return
      }
      deleteByGridKey(this.gridKey).then(res =>{
        if(res.status === 200){
          alert("删除成功")
          this.gridKey = ''
          this.getList()
        }else {
          alert("删除失败")
        }
      })
    },
    reInitLuckysheet(){
      this.sheet.destroy()
      this.createSheet('luckysheet')
    }
  },
  watch:{
    gridKey:function (newV){
      if(newV){
        this.getDetail(()=>{
          this.reInitLuckysheet()
        })
      }else {
        this.reInitLuckysheet()
      }


    }
  },
  created() {
    this.getList()
  },
  mounted() {
    this.createSheet('luckysheet')
  }

}
</script>

<style scoped>
.workList{
  float: left;
  width: 15%
}
.workArea {
  float: right;
  width: 85%
}
.workList li{
  text-align: left;
  cursor: pointer;
}
.workArea > .toolbar{
  margin-bottom: 10px;
  text-align: left;
}
.workArea > .toolbar > button{
  margin-right: 10px;
}

.clicked {
  font-weight: bold;
}
</style>
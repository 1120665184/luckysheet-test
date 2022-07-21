import axios from 'axios'

const service = axios.create({
    baseURL : '/luckysheet-service',
    timeout : 10000
})
service.interceptors.request.use(config => config,error => Promise.reject(error))
service.interceptors.response.use(response => response , error => Promise.reject(error))

export default service
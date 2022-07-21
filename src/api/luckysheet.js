import request from '@/utils/request'

/**
 * 整体保存或修改
 * @param data
 * @returns {AxiosPromise}
 */
export function saveOrEdit(data){
    return request({
        url : '/luckysheet',
        method: 'post',
        data
    })
}

/**
 * 获取工作簿列表
 * @param params
 * @returns {AxiosPromise}
 */
export function getList(params){
    return request({
        url : '/luckysheet',
        method: 'get',
        params
    })
}

/**
 * 获取基本信息
 * @param gridKey
 * @returns {AxiosPromise}
 */
export function findDetail(gridKey){
    return request({
        url:`/luckysheet/${gridKey}`,
        method:'get'
    })
}

/**
 * 删除
 * @param gridKey
 * @returns {AxiosPromise}
 */
export function deleteByGridKey(gridKey){
    return request({
        url:'/luckysheet',
        method:'delete',
        data:[gridKey]
    })
}
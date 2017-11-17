<?php
require_once('Classes/Application.class.php');


/*功能灵活导出Excel数据
   *@param 关键字传参 参数说明 array('fileName'=>'','lineWidth'=>array(),'headArr'=>array(),'bodyArr'=>array(),'fontColor'=>array())
   * fileName string 文件名字 后缀名 .xls 必须
   * lineWidth array 列间距 eg  array('A'=>'20','B'=>'30'.....) 非必需
   *
   * headArr array 首行标题 eg  array('序号',''.......)  非必需
   *bodyArr array 导出的数据 eg array(array('1','2','3',......))必须
   *
   * */

$app_model = new Appplication_class();





$data['fileName'] = 'hhhh.xls';
$data['headArr'] = ['数字','数字'];
$data['bodyArr'] = [[1,2,3],[1,2,3],[1,2,3]];
$app_model->dumpExcel($data);
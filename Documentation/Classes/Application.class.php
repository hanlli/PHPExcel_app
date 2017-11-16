<?php
/*
 * @author hanli
 * @description PHPExcel 应用封装
 * excel的简单封装 带有简单导出可以灵活设置表头宽度，内容颜色，
 */
class  Appplication_class{
    public function __construct()
    {
        parent::__construct();
        error_reporting(E_ALL);
        ini_set('display_errors', TRUE);
        ini_set('display_startup_errors', TRUE);

    }
    /*功能灵活导出Excel数据
     *@param 关键字传参 参数说明 array('fileName'=>'','lineWidth'=>array(),'headArr'=>array(),'bodyArr'=>array(),'fontColor'=>array())
     * fileName string 文件名字 后缀名 .xls 必须
     * lineWidth array 列间距 eg  array('A'=>'20','B'=>'30'.....) 非必需
     *
     * headArr array 首行标题 eg  array('序号',''.......)  非必需
     *bodyArr array 导出的数据 eg array(array('1','2','3',......))必须
     *
     * */
    public function dumpExcel()
    {
        $fun_arr = func_get_args()[0];
        if(empty($fun_arr)) return;


        //引入PHPExcel对象
        require_once('PHPExcel.php');
        require_once('PHPExcel/IOFactory.php');

        //创建PHPExcel对象
        $objPHPExcel = new PHPExcel();

        $objPHPExcel->getProperties();
        $objActSheet = $objPHPExcel->getActiveSheet();
        if(!isset($fun_arr['headArr']) || !isset($fun_arr['bodyArr'])) return;

        if(array_key_exists('headArr',$fun_arr)){
            if(!is_array($fun_arr['headArr']))return;
            $span = ord("A");
            foreach ($fun_arr['headArr'] as  $item){
                $j = chr($span);
                $objActSheet->setCellvalue($j.'1',$item);
                if(isset($fun_arr['lineWidth']) && array_key_exists($j,$fun_arr['lineWidth'])){
                    $objActSheet->getColumnDimension($j)->setWidth($fun_arr['lineWidth'][$j]);
                }else{
                    $objActSheet->getColumnDimension($j)->setWidth('20');
                }
                $span++;
            }
        }
        if(array_key_exists('bodyArr',$fun_arr)){
            if(!is_array($fun_arr['bodyArr']))return;
            $i = 2;
            foreach ($fun_arr['bodyArr'] as $hineValue){
                $span = ord("A");
                foreach ($hineValue as $lineValue){
                    $j = chr($span);
                    if(!empty($fun_arr['fontColor'])){

                        if(array_key_exists($j,$fun_arr['fontColor']) ){
                            $objRichText2 = new PHPExcel_RichText();
                            $objRichText2->createText("");
                            $objRed = $objRichText2->createTextRun($lineValue);
                            if(array_key_exists($lineValue,$fun_arr['fontColor'][$j])){
                                $objRed->getFont()->setColor( new PHPExcel_Style_Color( 'FF'.$fun_arr['fontColor'][$j][$lineValue]) );

                            }else{
                                $objRed->getFont()->setColor( new PHPExcel_Style_Color( 'FF'.$fun_arr['fontColor'][$j][$lineValue]) );
                            }
                            $objPHPExcel->getActiveSheet()->getCell($j .$i)->setValue($objRichText2);
                            $objPHPExcel->getActiveSheet()->getStyle($j .$i)->getAlignment()->setWrapText(true);
                        }else{
                            $objActSheet->setCellvalue($j.$i,$lineValue);
                        }
                    }else{
                        $objActSheet->setCellvalue($j.$i,$lineValue);
                    }


                    $span++;
                }
                $i++;
            }
        }

        //设置活动单指数到第一个表,所以Excel打开这是第一个表
        $objPHPExcel->setActiveSheetIndex(0);
        ob_end_clean(); //清除缓冲区,避免乱码
        if(array_key_exists('fileName',$fun_arr)){
            if(!is_string($fun_arr['fileName']))return;
            $fileName = iconv("utf-8", "gb2312", $fun_arr['fileName']);
        }else{
            return;
        }
        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition: attachment;filename=\"$fileName\"");
        header('Cache-Control: max-age=0');

        $objWriter = IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output'); //文件通过浏览器下载

        exit();
    }
    /*功能灵活导出Excel数据
    *@param 关键字传参 参数说明 array('fileName'=>'','lineWidth'=>array(),'headArr'=>array(),'bodyArr'=>array(),'fontColor'=>array())
    * fileName string 文件名字 后缀名 .xls 必须
    * lineWidth array 列间距 eg  array('A'=>'20','B'=>'30'.....) 非必需
    *
    * headArr array 首行标题 eg  array('序号',''.......)  非必需
    *bodyArr array 导出的数据 eg array(array('1','2','3',......))必须
    *
    * */
    public function dumpExcels()
    {
        $fun_arr = func_get_args()[0];
        if(empty($fun_arr)) return;


        //引入PHPExcel对象
        require_once('PHPExcel.php');
        require_once('PHPExcel/IOFactory.php');

        //创建PHPExcel对象
        $objPHPExcel = new PHPExcel();

        $objPHPExcel->getProperties();
        $objActSheet = $objPHPExcel->getActiveSheet();
        if(!isset($fun_arr['headArr']) || !isset($fun_arr['bodyArr'])) return;

        if(array_key_exists('headArr',$fun_arr)){
            if(!is_array($fun_arr['headArr']))return;
            $span = ord("A");
            foreach ($fun_arr['headArr'] as  $item){
                $j = chr($span);
                $objActSheet->setCellvalue($j.'1',$item);
                if(isset($fun_arr['lineWidth']) && array_key_exists($j,$fun_arr['lineWidth'])){
                    $objActSheet->getColumnDimension($j)->setWidth($fun_arr['lineWidth'][$j]);
                }else{
                    $objActSheet->getColumnDimension($j)->setWidth('20');
                }
                $span++;
            }
        }
        if(array_key_exists('bodyArr',$fun_arr)){
            if(!is_array($fun_arr['bodyArr']))return;
            $i = 2;
            foreach ($fun_arr['bodyArr'] as $hineValue){
                $span = ord("A");
                foreach ($hineValue as $lineValue){
                    $j = chr($span);
                    if(!empty($fun_arr['fontColor'])){

                        if(array_key_exists($j,$fun_arr['fontColor']) ){
                            $objRichText2 = new PHPExcel_RichText();
                            $objRichText2->createText("");
                            $objRed = $objRichText2->createTextRun($lineValue);
                            if(array_key_exists($lineValue,$fun_arr['fontColor'][$j])){
                                $objRed->getFont()->setColor( new PHPExcel_Style_Color( 'FF'.$fun_arr['fontColor'][$j][$lineValue]) );

                            }else{
                                $objRed->getFont()->setColor( new PHPExcel_Style_Color( 'FF'.$fun_arr['fontColor'][$j][$lineValue]) );
                            }
                            $objPHPExcel->getActiveSheet()->getCell($j .$i)->setValue($objRichText2);
                            $objPHPExcel->getActiveSheet()->getStyle($j .$i)->getAlignment()->setWrapText(true);
                        }else{
                            $objActSheet->setCellvalue($j.$i,$lineValue);
                        }
                    }else{
                        $objActSheet->setCellvalue($j.$i,$lineValue);
                    }


                    $span++;
                }
                $i++;
            }
        }

        //设置活动单指数到第一个表,所以Excel打开这是第一个表
        $objPHPExcel->setActiveSheetIndex(0);
        ob_end_clean(); //清除缓冲区,避免乱码
        if(array_key_exists('fileName',$fun_arr)){
            if(!is_string($fun_arr['fileName']))return;
            $fileName = iconv("utf-8", "gb2312", $fun_arr['fileName']);
        }else{
            return;
        }
        $save_path = Wxconfig::LOGPATH.'xls/'.$fileName;
        //  $save_path = iconv("utf-8", "gb2312", $save_path);
        $objWriter = IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save($save_path); //保存本地


    }

    public function dumpStyleExcel($filename ='test.xlsx' ,$title = '****月积分报表',$body = null,$formstyle_border = null,$formstyle_body = null){
        require_once('PHPExcel.php');
        require_once('PHPExcel/IOFactory.php');

        if($body == null){
            return;
        }
        $objPHPExcel = new PHPExcel();
        $objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
            ->setLastModifiedBy("Maarten Balliauw")
            ->setTitle("Office 2007 XLSX Test Document")
            ->setSubject("Office 2007 XLSX Test Document")
            ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
            ->setKeywords("office 2007 openxml php")
            ->setCategory("Test result file");



        $objPHPExcel->setActiveSheetIndex(0);

        $sharedStyle1 = new PHPExcel_Style();
        $sharedStyle2 = new PHPExcel_Style();

        $sharedStyle1->applyFromArray(
            array('fill' 	=> array(
                'type'		=> PHPExcel_Style_Fill::FILL_SOLID,
                'color'		=> array('argb' => 'FFCCFFCC')
            ),
                'borders' => array(
                    'bottom'	=> array('style' => PHPExcel_Style_Border::BORDER_THIN),
                    'right'		=> array('style' => PHPExcel_Style_Border::BORDER_MEDIUM)
                )
            ));

        $sharedStyle2->applyFromArray(
            array('fill' 	=> array(
                'type'		=> PHPExcel_Style_Fill::FILL_SOLID,
                'color'		=> array('argb' => 'FFFFFF00')
            ),
                'borders' => array(
                    'bottom'	=> array('style' => PHPExcel_Style_Border::BORDER_THIN),
                    'right'		=> array('style' => PHPExcel_Style_Border::BORDER_MEDIUM)
                )
            ));
        if($formstyle_border == null){
            $objPHPExcel->getActiveSheet()->setSharedStyle($sharedStyle1, "A1:H40");
        }else{
            $objPHPExcel->getActiveSheet()->setSharedStyle($sharedStyle1, $formstyle_border);
        }
        if($formstyle_body == null){
            $objPHPExcel->getActiveSheet()->setSharedStyle($sharedStyle2, "C3:F38");
        }else{
            $objPHPExcel->getActiveSheet()->setSharedStyle($sharedStyle2, $formstyle_body);
        }
        $objActSheet = $objPHPExcel->getActiveSheet();
        $objActSheet->getColumnDimension('C')->setWidth('20');
        $objActSheet->getColumnDimension('D')->setWidth('20');
        $objActSheet->getColumnDimension('E')->setWidth('20');
        $objActSheet->getColumnDimension('F')->setWidth('20');
        $objActSheet->getColumnDimension('G')->setWidth('20');
        $objActSheet->getColumnDimension('H')->setWidth('20');
        $objActSheet->getColumnDimension('I')->setWidth('20');

        $objPHPExcel->getActiveSheet()->mergeCells('C1:I1');
        $objActSheet->setCellvalue('C1',$title);

        $objPHPExcel->getActiveSheet()->getStyle('C1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $i = 2;

        foreach($body as $value){
            $j = ord('A');
            foreach($value as $val){

                $objActSheet->setCellvalue(chr($j).$i,$val);
                $j++;
            }
            $i++;

        }

        //设置活动单指数到第一个表,所以Excel打开这是第一个表
        $objPHPExcel->setActiveSheetIndex(0);
        ob_end_clean(); //清除缓冲区,避免乱码

        $fileName = iconv("utf-8", "gb2312", $filename);

        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition: attachment;filename=\"$fileName\"");
        header('Cache-Control: max-age=0');

        $objWriter = IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output'); //保存本地

        exit();

    }



}
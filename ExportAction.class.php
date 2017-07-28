<?php
/**
 * 对象转换操作
 * User: noah
 * Date: 2016/8/12  15:13
 */

namespace Main\Action;

use Main\Common\FileUtil;
use Main\Model\ExcelConfigModel;
use Main\Model\ExcelConfigSUBModel;

require "./phpword/PHPWord.php";
require "./phpexcel/PHPExcel.php";

/**
 * 通用报告生成类
 * Class ReportUtil
 * @package Main\Common
 */
class ExportAction
{
    /**
     * 通过模板创建的word的通用方法
     * @param $templateName 模板名称。temple.docx
     * @param $data  模板中对应的数据
     * $table1 = array(array("no"=>123,"name"=>"abc"),array("no"=>345,"name"=>"qwe"),array("no"=>665,"name"=>"kug"));
     * $table2 = array(array("no"=>111,"name"=>"abc"),array("no"=>222,"name"=>"qwe"),array("no"=>665,"name"=>"kug"));
     * $obj = array("no"=>"#001","name"=>"赵首重","age"=>"100","tel"=>"13800808");
     * $table = array("t1"=>$table1,"t2"=>$table2);
     * $data["obj"] = $obj ;
     * $data["table"] = $table;
     * @param $fileName 生成后的文件名称 newword.docx 默认guid
     * @return bool
     * 更多例子请查看ExportController.php
     */
    public function createWord($templateName,$data,$fileName){
        $template = $this->getTemplatePath().'/'.$templateName;
        if(!file_exists($template))return false;
        $PHPWord = new \PHPWord();
        $document = $PHPWord->loadTemplate($template);
        foreach ($data as $key => $value){
            //图片
            if($key=="img"){
                $document->setImg($value);
            }else {
                if (is_array($value)) {
                    foreach ($value as $k => $v) {
                        //Table
                        if (is_array($v)) {
                            $s = $this->getTableKey($v);
                            $tableKey = $k . '.' . $s;
                            //var_dump($tableKey.".........");
                            $table = $document->getTable($tableKey);
                            foreach ($v as $p => $d) {
                                $document->addRow($table, $tableKey, $d);
                            }
                        } else {
                            $objKey = $key . '.' . $k;
                            //var_dump($signKey,$v);//"${obj.no}" string(4) "#001"
                            $document->setValue($objKey, $v);
                        }
                    }
                }
            }
        }
        //* 生成后的文件存放路径 默认项目目录下tmp/日期
        $filePath = $this->getTMPPath($fileName,"docx");
        $document->save($filePath);
        return $filePath;
    }

    private function getTableKey($table){
        $sign = '';
        //取第一个属性做为key
        if(is_array($table)){
            foreach ($table as $key =>$value){
                if(is_array($value)){
                    foreach ($value as $k =>$v){
                        $sign = $k;
                        break;
                    }
                }
                break;
            }
        }
        return $sign;
    }

    private function GUID()
    {
        return create_guid();
    }

    public function getPath(){
        //D:\dev\php\taskms
        $root = realpath(str_replace("ThinkPHP","".THINK_PATH));
        return $root;
    }
    
    public function getTemplatePath(){
        $root = $this->getPath();
        return $root.'/'.C("TEMPLATE_TEMP_PATH");
    }

    public function getTMP(){
        $root = $this->getPath();
        return $root.'/'.C("FILE_TEMP_PATH");
    }

    /**
     * 返回导出的文件存放的临时文件夹目录和名称
     * D:\dev\taskms\1223455.dox
     * @param $fileName 文件名
     * @param $postfix  文件后缀
     * @return string
     */
    public function getTMPPath($fileName,$postfix){
        $filePath = $this->getTMP().'/'.date("Ymd");
        if (!file_exists($filePath)){
            FileUtil::createDir($filePath);
        }
        if(!$fileName){
            $fileName = $this->GUID();
        }
        return $filePath.'/'.$fileName.'.'.$postfix;
    }

    /**
     * 导出excel，默认的方式是向下扩展行
     * @param $data array("table"=>array(0=>array("col1"=>11,"col2"=>22),1=>array("col1"=>33,"col2"=>44)),"form"=>array("form1"=>"aa","form2"=>"bb"))
     * 如果向X轴方向扩展，数据格式稍有差异
     * @param $configId
     * @param $fileName
     * @return string|void
     * 更多例子请查看ExportController.php
     */
    public function createExcel($data,$configId,$fileName){
        $configM = new ExcelConfigModel();
        $config = $configM->getConfigById($configId);
        if(!$config){ return; }
        $configSUBM = new ExcelConfigSUBModel();
        $configSUB = $configSUBM->getConfigSUBByConfigId($configId);

        $templatePath = $this->getTemplatePath()."/".$config["template_name"];

        $objPHPExcel = \PHPExcel_IOFactory::load($templatePath);
        $baseCol = getStringChar(strtoupper($config["begin_row"]));//开始列数
        $baseRow = intval(getStringNum($config["begin_row"]));//开始行数
        $sheet = $objPHPExcel->getActiveSheet();
        //处理循环休表格
        if($data["table"]){
            $vector = strtoupper($config["vector"]);
            $d = $data["table"];
            //向X轴方向扩展
            if(strpos($vector,"X")!==false){
                $this->createTableX($sheet,$baseRow,$baseCol,$d);
            }else{
                //向Y轴方向扩展
                if(!$configSUB){ return ; }

                $this->createTableY($sheet,$baseRow,$d,$configSUB);
            }
            $sheet->removeRow($baseRow,1);
        }
        //处理非循环休部分
        if($data["form"]){
            $d = $data["form"];
            $this->addForm($sheet,$d,$configSUB);
        }
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');

        $filePath = $this->getTMPPath($fileName,"xls");
        $objWriter->save($filePath);
        return $filePath;
    }

    public function createTableY($sheet,$baseRow,$data,$configSUB){
        foreach($data as $r => $dataRow) {
            $row = $baseRow+1 + $r;
            $sheet->insertNewRowBefore($row,1);
            $this->addRow($sheet,$row,$dataRow,$configSUB);
        }
    }

    public function createTableX($sheet,$baseRow,$baseCol,$data){
        if(count($data)>1){
            $colTitle = array_keys($data[0]);//获取列标题,要求每一行的列数相等
            foreach($data as $r => $dataRow) {
                $row = $baseRow+1 + $r;
                $sheet->insertNewRowBefore($row,1);
                $this->addColumn($sheet,$baseCol,$row,$dataRow,$colTitle);
            }
        }
    }

    private function addRow($sheet,$row,$data,$config){
        foreach ($config as $c){
            if($c["is_form"]==1){ continue; }
            $key = $c["col"];
            $d = $data[$key];
            $position = $c["position"].$row;
            $expression = $c["expression"];
            if($expression){
                $e = str_replace('#var',"\"$d\"",$expression);
                $e = str_replace('#row',$row,$e);
                $sheet->setCellValue($position, '='.$e);
            }else{
                if($d){
                    $sheet->setCellValue($position,$d);
                }
            }
        }
    }

    /**
     * @param $sheet
     * @param $baseCol  开始的列号如F，G等
     * @param $currentRow  当前行
     * @param $dataRow
     * @param $colTitle
     */
    public function addColumn($sheet,$baseCol,$currentRow,$dataRow,$colTitle){
        $beginCol = $baseCol;
        foreach ($colTitle as $c=>$value){
            $d = $dataRow[$value];
            $position = $beginCol;
            $sheet->setCellValue($position.$currentRow,$d);
            $beginCol = $this->getNextCol($beginCol);
        }
    }

    public function addForm($sheet,$data,$config){
        foreach ($config as $c) {
            if ($c["is_form"] == 0) {
                continue;
            }
            $key = $c["col"];
            $d = $data[$key];
            $position = $c["position"];
            $sheet->setCellValue($position,$d);
        }
    }

    /**
     * 获取指定列序号的下一列序号
     * 如指定A列，下一列为B，指定Z列，下一列为AA
     * @param $baseCol
     * @return null
     */
    private function getNextCol($baseCol){
        if (empty($baseCol)) return null;
        $len = mb_strlen($baseCol);
        $charArray = array();
        for ($i = 0; $i < $len; $i++) {
            $charArray[] = mb_substr($baseCol, $i, 1);
        }
        $char = array();
        $index = count($charArray);
        $nextAdd = false;//前一位是否自增
        for ($i=$index-1;$i>=0;$i--){
            $currentChar = $charArray[$i];
            $charnum = ord($currentChar);
            if($nextAdd||$i==$index-1){
                $nextAdd=false;
                $charnum++;
            }
            //从最后一个字符开始 是否已经是最后一个字母Z
            if($charnum>90){
                $nextAdd = true;
                $char[$i]="A";
            }else{
                $s = chr($charnum);
                array_unshift($char,$s);
            }
        }
        if($nextAdd){
            array_unshift($char,"A");
        }
        return implode($char);
    }

    public function dashboard($data,$fileName){

        $templatePath = $this->getTMPPath($fileName,"xlsx");
        //处理数据
        $data = $this->formatData($data);

        $objPHPExcel = new \PHPExcel();
        $objWorksheet = $objPHPExcel->getActiveSheet();
        //设置第一列的宽度
        $objWorksheet->getColumnDimension()->setWidth(20);
        $objWorksheet->fromArray($data);

        $lastColumn = $this->getLastColumn($data);
        //echo $lastColumn;
        //柱状图
        $columnChart = $this->getColumnChart($lastColumn);
        //线图
        $lineChart = $this->getLineChart($lastColumn);
        //样式
        $layout = new \PHPExcel_Chart_Layout();
        $layout->setShowVal(true);
        $layout->setShowPercent(false);

        $plotArea = new \PHPExcel_Chart_PlotArea($layout, array($columnChart,$lineChart));

        $legend = new \PHPExcel_Chart_Legend(\PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);

        $title = new \PHPExcel_Chart_Title('Test Chart');

        $chart = new \PHPExcel_Chart('chart1',$title,$legend,$plotArea,true,0,NULL,NULL);

        //图表的位置
        $chart->setTopLeftPosition('A10');
        $chart->setBottomRightPosition('M50');
        $objWorksheet->addChart($chart);

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->setIncludeCharts(TRUE);
        $objWriter->save($templatePath);
        return $templatePath;
    }

    /**
     * 把数值0处理为字符串0
     * @param $data
     * @return mixed
     */
    private function formatData($data){
        for ($i=0;$i<count($data);$i++){
            $row = $data[$i];
            for ($j=0;$j<count($row);$j++){
                $value = $row[$j];
                if($value===0){
                    $data[$i][$j] = "0";
                }
            }
        }
        return $data;
    }
    /**
     * 获取最后一列序号
     * @param $data
     * @return mixed|string
     */
    private function getLastColumn($data){
        if(count($data)<=0){
            return "A";
        }
        $colA = $data[0];
        $colNum = count($colA);
        $colNo = $this->getColumnNUM($colNum-1);
        return $colNo;
    }

    /**
     * 获取指定值对应的列序号
     * @param int $pColumnIndex
     * @return mixed
     */
    private function getColumnNUM($pColumnIndex = 0)
    {
        $_indexCache = array();

        if (!isset($_indexCache[$pColumnIndex])) {
            // Determine column string
            if ($pColumnIndex < 26) {
                $_indexCache[$pColumnIndex] = chr(65 + $pColumnIndex);
            } elseif ($pColumnIndex < 702) {
                $_indexCache[$pColumnIndex] = chr(64 + ($pColumnIndex / 26)) . chr(65 + $pColumnIndex % 26);
            } else {
                $_indexCache[$pColumnIndex] = chr(64 + (($pColumnIndex - 26) / 676)) . chr(65 + ((($pColumnIndex - 26) % 676) / 26)) . chr(65 + $pColumnIndex % 26);
            }
        }
        return $_indexCache[$pColumnIndex];
    }

    /**
     * 获取柱状图
     * @param $lastColumn
     * @return \PHPExcel_Chart_DataSeries
     */
    private function getColumnChart($lastColumn){
        $dataSeriesValues = $this->getColumnChartData($lastColumn);
        $dataSeriesLabels = $this->getColumnChartLabels();
        $xAxisTickValues = $this->getColumnChartValues($lastColumn);
        $series = new \PHPExcel_Chart_DataSeries(
            \PHPExcel_Chart_DataSeries::TYPE_BARCHART,		// plotType
            \PHPExcel_Chart_DataSeries::GROUPING_STACKED,	// plotGrouping
            range(0, count($dataSeriesValues)-1),			// plotOrder
            $dataSeriesLabels,								// plotLabel
            $xAxisTickValues,								// plotCategory
            $dataSeriesValues								// plotValues
        );
        $series->setPlotDirection(\PHPExcel_Chart_DataSeries::DIRECTION_COL);
        return $series;
    }

    /**
     * 获取柱状图的列标签名
     * @return array
     */
    private function getColumnChartLabels(){
        //柱状图
        $dataSeriesLabels = array(
            new \PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$4', NULL, 1),//	Pass
            new \PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$5', NULL, 1),//	Pass with DA
            new \PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$6', NULL, 1),//	Inprocess
            new \PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$7', NULL, 1),//	Evaluation
            new \PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$8', NULL, 1) //	Not Pass
        );
        return $dataSeriesLabels;
    }

    /**
     * 获取柱状图的X轴方向的值的区域
     * @return array
     */
    private function getColumnChartValues($lastColumn){
        $xAxisTickValues = array(
            new \PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$1:$'.$lastColumn.'$1', NULL, 4)	//wk1 to wk4
        );
        return $xAxisTickValues;
    }

    /**
     * 获取柱状图的数据区域
     * @param $lastColumn
     * @return array
     */
    private function getColumnChartData($lastColumn){
        $dataSeriesValues = array(
            new \PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$4:$'.$lastColumn.'$4', NULL, 4),
            new \PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$5:$'.$lastColumn.'$5', NULL, 4),
            new \PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$6:$'.$lastColumn.'$6', NULL, 4),
            new \PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$7:$'.$lastColumn.'$7', NULL, 4),
            new \PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$8:$'.$lastColumn.'$8', NULL, 4)
        );
        return $dataSeriesValues;
    }

    /**
     * 获取线图
     * @param $lastColumn 有数据的最后一个列的列序号 如D，AB
     * @return \PHPExcel_Chart_DataSeries
     */
    private function getLineChart($lastColumn){
        $dataSeriesValues = $this->getLineChartData($lastColumn);
        $dataSeriesLabels = $this->getLineChartLabels();
        $series = new \PHPExcel_Chart_DataSeries(
            \PHPExcel_Chart_DataSeries::TYPE_LINECHART,
            \PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
            range(0,1),
            $dataSeriesLabels,
            null,
            $dataSeriesValues
        );
        return $series;
    }

    /**
     * 获取线图标签名
     * @return array
     */
    private function getLineChartLabels(){
        $dataSeriesLabels = array(
            new \PHPExcel_Chart_DataSeriesValues('String','Worksheet!$A$2',null,1),//Original Plan
            new \PHPExcel_Chart_DataSeriesValues('String','Worksheet!$A$3',null,1),//Plan
        );
        return $dataSeriesLabels;
    }

    /**
     * 获取线图数据区域
     * @param $lastColumn 最后一个列
     * @return array
     */
    private function getLineChartData($lastColumn){
        $dataSeriesValues = array(
            new \PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$2:$'.$lastColumn.'$2', NULL, 4),
            new \PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$3:$'.$lastColumn.'$3', NULL, 4)
        );
        return $dataSeriesValues;
    }
}
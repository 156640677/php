<?php
namespace Main\Controller;
use Main\Action\ExportAction;
use Main\Action\GNUMAction;
use Main\Service\FixtureService;
use Think\Controller;
use Think\Model;

class ExportController extends Controller {
    
    public function Index(){
        $this->display('Export:index');
    }

    public function word(){
        $table1 = array(array("no"=>123,"name"=>"abc"),array("no"=>345,"name"=>"qwe"),array("no"=>665,"name"=>"kug"));
        $table2 = array(array("no"=>111,"name"=>"abc"),array("no"=>222,"name"=>"qwe"),array("no"=>665,"name"=>"kug"));
        $obj = array("no"=>"#001","name"=>"张三","age"=>"100","tel"=>"13800808");
        $table = array("t1"=>$table1,"t2"=>$table2);
        $data["obj"] = $obj ;//这里的key可以为任意值
        $data["table"] = $table;//这里的key可以为任意值
        $export = new ExportAction();
        $root = $export->getPath().'/'.C("TEMPLATE_TEMP_PATH");
        $a001 = array("img1"=>array(
                            'img' => $root.'/image/001.jpg',//绝对路径，图片格式为jpg
                            'size' => array(100, 50)
                        ),
                        "img2"=>array(
                            'img' => $root.'/image/002.jpg',//绝对路径，图片格式为jpg
                            'size' => array(100, 50)
            ));
        $data["img"] =  $a001;//这里的key为固定值img

        //$root = realpath(str_replace("ThinkPHP","".THINK_PATH));
        //$template = $root.'/'.C("FILE_TEMP_PATH");
        //echo $root;//$template;
        $result = $export->createWord("test2.docx",$data,"test3.docx");
        echo $result;
    }

    public function excel(){
        /*$table1["table"] = array(array("col1"=>123,"col2"=>"2","col3"=>"aaa"),
                        array("col1"=>345,"col2"=>"3","col3"=>"bbb"),
                        array("col1"=>567,"col2"=>"4","col3"=>"ccc"));*/
        //
        /*$table2["table"] = array(array("col1"=>123,"col2"=>"2","col3"=>"aaa","col7"=>"3","col8"=>"6","col9"=>"2","col10"=>"4"),
                        array("col1"=>345,"col2"=>"3","col3"=>"bbb","col7"=>"2","col8"=>"8","col9"=>"6","col10"=>"6"),
                        array("col1"=>567,"col2"=>"4","col3"=>"ccc","col7"=>"2","col8"=>"3","col9"=>"3","col10"=>"4"));
        $table2["form"]=array("form1"=>"2017/4/8","form2"=>"2017/5/8");*/
        //向X轴方向扩展 该扩展不用在excel_sub表中配置列信息
        $table3["table"] = array(array("col1"=>123,"col2"=>"2","col3"=>"aaa"),
            array("col1"=>345,"col2"=>"3","col3"=>"bbb"),
            array("col1"=>567,"col2"=>"4","col3"=>"ccc"));

        $action = new ExportAction();
        $action->createExcel($table3,2);
    }

    public function dashboard(){
        //dashboard
        $data = array(
                    array('Week',	"wk1",	"wk2",	"wk3"),
                    array('Original Plan',   '12',   '15',		'21'),
                    array('Plan',   '10',   '73',		'86'),
                    array('Pass',   '10',   '11',		'19'),
                    array('Pass with DA',   "0",   '32',		'10'),
                    array('Inprocess',   '20',   '42',		'3'),
                    array('Evaluation',   '30',   '12',		'10'),
                    array('Not Pass',   "15",   "0",		"0")
                );
        $action = new ExportAction();
        $action->dashboard($data);
    }

}

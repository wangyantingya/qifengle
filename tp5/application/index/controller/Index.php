<?php
namespace app\index\controller;

use think\Loader;
use think\Controller;
use think\Db;
use PHPMailer\PHPMailer\PHPMailer;

class Index extends Controller
{
    public function index()
    {
        return "<a href='".url('daochu')."'>导出</a>";
    }
    public function excel()
    {
        $path = dirname(__FILE__); //找到当前脚本所在路径
        Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
        $PHPExcel = new \PHPExcel(); //实例化
        $iclasslist=db('iclass')->select();
        foreach($iclasslist as $key=> $v){
            $PHPExcel->createSheet();
            $PHPExcel->setactivesheetindex($key);
            $PHPSheet = $PHPExcel->getActiveSheet();
            $PHPSheet->setTitle($v['classname']); //给当前活动sheet设置名称
            $PHPSheet->setCellValue("A1", "编号")
                     ->setCellValue("B1", "姓名")
                     ->setCellValue("C1", "性别")
                     ->setCellValue("D1", "身份证号")
                     ->setCellValue("E1", "宿舍编号")
                     ->setCellValue("F1", "班级");//表格数据
            $userlist=db('users')->where("iclass=".$v['id'])->select();
            //echo db('users')->getLastSql();
            $i=2;
            foreach($userlist as $t)
            {
                $PHPSheet->setCellValue("A".$i, $t['id'])
                         ->setCellValue("B".$i, $t['username'])
                        ->setCellValue("C".$i, $t['sex'])
                        ->setCellValue("D".$i, $t['idcate'])
                        ->setCellValue("E".$i, $t['dorm_id'])
                        ->setCellValue("F".$i, $t['iclass']);
                        //表格数据
                $i++;
            }

        }
       // exit;
        $PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, "Excel2007"); //创建生成的格式
        header('Content-Disposition: attachment;filename="学生列表'.time().'.xlsx"'); //下载下来的表格名
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
    }
     
    public function daochu()
    {
        $path= dirname(__FILE__);
        Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 
        $excel= new \PHPExcel(); //实例化
        $classes=db('iclass')->select(); //查询所有班级
        foreach ($classes as $key => $clas)
        {
            $excel->createSheet(); //创建sheet
            $excel->setactivesheetindex($key);
            $phpsheet=$excel->getActiveSheet();
            $phpsheet->setTitle($clas['classname']); //每个的sheet的名称

            $cellName = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AB');
            
            $lists=Db::query('SHOW FULL COLUMNS from wx_users');
                foreach($lists as $key=>$lis)
                {//获取当前表结构,赋给sheet的第一行
                   $comment=$lis['Comment']?$lis['Comment']:$lis['Field'];
                   $phpsheet->setCellValue($cellName[$key].'1',$comment);
                }
        // echo "<pre/>";
        // print_r($students);exit;
            $students=db('users')->where("iclass=".$clas['id'])->select();
             $i=2;
                foreach($students as $key=>$use)
                {
                    $j=0;
                    foreach($use as $u)
                    {
                        $phpsheet->setCellValue($cellName[$j].$i,$u);
                        $j++;
                    }
                    $i++;
                }
        
        }
        

        $PHPWriter = \PHPExcel_IOFactory::createWriter($excel, "Excel2007"); //创建生成的格式
        header('Content-Disposition: attachment;filename="学生列表'.'.xlsx"'); //下载下来的表格名
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件


    }
    //导入
    public function daoru()
    {
        return $this->fetch();
    }
    public function dao_dr()
    {
        Loader::import('PHPExcel.PHPExcel');
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory');
        Loader::import('PHPExcel.PHPExcel.Reader.Excel5');
        //获取表单上传文件
        $file = request()->file('excel');
 //        echo "<pre>";
 // print_r($file);exit;
        $info =$file->validate(['ext' => 'xlsx'])->move(ROOT_PATH .'public'.DS .'uploads');

        //上传验证后缀名,以及上传之后移动的地址
        if($info){
         //            echo $info->getFilename();
         $exclePath = $info->getSaveName();  //获取文件名
             $file_name = ROOT_PATH . 'public' . DS . 'uploads' . DS . $exclePath;   //上传文件的地址
             $objReader =\PHPExcel_IOFactory::createReader('Excel2007');
             $obj_PHPExcel =$objReader->load($file_name, $encode = 'utf-8');  //加载文件内容,编码utf-8
             //echo "<pre>";
             $excel_array=$obj_PHPExcel->getsheet(0)->toArray();   //转换为数组格式
             array_shift($excel_array);  //删除第一个数组(标题);
             $users = [];
             foreach($excel_array as $k=>$v){
                 $users[$k]['username'] = $v[0];
                 $users[$k]['sex'] = $v[1];
                 $users[$k]['idcate'] = $v[2];
                 $users[$k]['dorm_id'] = $v[3];
                 $users[$k]['adress'] = $v[4];
             }
             // print_r($users);exit;
             $dbs=Db::name('users')->insertAll($users); //批量插入数据
             if($dbs)
             {
                echo "ok";
             }else{
                echo "fail";
             }
         } else {
             echo $file->getError();

        }
    }
    // 邮箱
    public function reg()
    {
        $toemail=input('post.email');
        $username=input('post.username');
        $title="你好,".$username.'欢迎注册相亲网';
        $body="你好，".$username.'相亲网欢迎你的加入 ';
        sendmail($toemail,$title,$body);
    }
    public function zhuce()
    {
        return $this->fetch();
    }
}

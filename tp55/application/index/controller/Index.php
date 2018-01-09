<?php
namespace app\index\controller;

use think\Loader;
use think\Controller;
use think\Db;
use PHPMailer\PHPMailer\PHPMailer;
class Index extends Controller
{
    //引用PHPMailer拓展实现数据库导出excel
    public function index()
    {
        return "<a href='".url('excel')."'>导出</a>";
    }
    //导出--班级分类学生
    public function excel()
    {
        $path = dirname(__FILE__); //找到当前脚本所在路径
        Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
        $PHPExcel = new \PHPExcel(); //实例化
        $iclasslist=db('iclass')->select();
        //echo "<pre>";print_r($iclasslist);exit;
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
            //echo db('users')->getLastSql();exit;
        
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
        $PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, "Excel2007"); //创建生成的格式
        header('Content-Disposition: attachment;filename="学生列表'.time().'.xlsx"'); //下载下来的表格名
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
    }
    //导出--回家班级分类学生
    public function excel_home()
    {
        $path = dirname(__FILE__); //找到当前脚本所在路径
        Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
        $PHPExcel = new \PHPExcel(); //实例化
        $iclasslist=db('iclass')->select();
        //echo "<pre>";print_r($iclasslist);exit;
        foreach($iclasslist as $key=> $v){
            $PHPExcel->createSheet();
            $PHPExcel->setactivesheetindex($key);
            $PHPSheet = $PHPExcel->getActiveSheet();
            $PHPSheet->setTitle($v['classname']); //给当前活动sheet设置名称
            $userlist=db('backhome')->where('iclass='.$v['id'])->select();
            // echo "<pre>";print_r($userlist);exit;
            $column=Db::query('SHOW FULL COLUMNS from wx_backhome');
           // echo "<pre>";print_r($column);exit;
            $arr=array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
            foreach($column as $key=>$v){
                $comment=$v['Comment']?$v['Comment']:$v['Field'];
                $PHPSheet->setCellValue($arr[$key].'1',$comment);
            }
        
            $i=2;
            foreach($userlist as $ulist)
            {
                $j=0;
                foreach($ulist as $key => $v){
                    $PHPSheet->setCellValue($arr[$j].$i,$v);
                    $j++;
                }
                $i++;
            }

        }
     
        $PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, "Excel2007"); //创建生成的格式
        header('Content-Disposition: attachment;filename="周末回家申请'.time().'.xlsx"'); //下载下来的表格名
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
    }
    //引用PHPMailer拓展实现excel导入数据库
    public function insemail()
    {
       return  $this->fetch();
    }
    //导入
    public function inserexcel()
    {
        Loader::import('PHPExcel.PHPExcel');
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory');
        Loader::import('PHPExcel.PHPExcel.Reader.Excel5');
        //获取表单上传文件
           $file = request()->file('excel');
           $info = $file->validate(['ext' => 'xlsx'])->move(ROOT_PATH . 'public' . DS . 'uploads');

        if($info) {
          // echo $info->getFilename();
            $exclePath = $info->getSaveName();  //获取文件名
            $file_name = ROOT_PATH . 'public' . DS . 'uploads' . DS . $exclePath;   //上传文件的地址
            $objReader =\PHPExcel_IOFactory::createReader("Excel2007");
            $obj_PHPExcel =$objReader->load($file_name, $encode = 'utf-8');  //加载文件内容,编码utf-8
            echo "<pre>";
            $excel_array=$obj_PHPExcel->getsheet(0)->toArray();   //转换为数组格式
            array_shift($excel_array);  //删除第一个数组(标题);
            $users = [];
            foreach($excel_array as $k=>$v) {
                $users[$k]['username'] = $v[0];
                $users[$k]['sex'] = $v[1];
                $users[$k]['idcate'] = $v[2];
                $users[$k]['dorm_id'] = $v[3];
                $users[$k]['iclass'] = $v[4];
                // print_r($users) ;exit;

            }
            //echo "<pre>";print_r($users);exit;
           $add=Db::name('users')->insertAll($users); //批量插入数据
           if($add){
            $this->success('succ');
           }else{
            $this->error('fail');
           }
        } else {
            echo $file->getError();
        }
    }
    public function wyeamil()
    {
        return $this->fetch();
    }
    public function doeamil()
    {
        $email=input('post.email');
        $username=input('post.username');
        $title="你好".$username."欢迎注册";
        $body="欢迎加入慕课网".$email.'欢迎你的加入,现在是我们的一员';
        sendmailer($email,$title,$body);
    }
    
}

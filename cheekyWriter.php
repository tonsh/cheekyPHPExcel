<?php
class CheekyWriter{
    private $fileType = 'xls';
    private $fileName = 'default.xls';
    private $template = null; //是否写入已有excel文件，null新建Excel
    private $isLimitMemory = false;
    private $isLimitTime = false;
    public $excelObj = null;

    public function __construct(array $setting = array()) {
        $attributs = array('fileType', 'fileName', 'template',
            'isLimitMemory', 'isLimitTime');
        foreach($attributs as $attr) {
            $setter = "set".ucfirst($attr);
            if(isset($setting[$attr])) $this->$setter($setting[$attr]);
        }

        include_once('./phpExcel/PHPExcel.php');
        if($this->template) {
            $this->excelObj = PHPExcel_IOFactory::load($this->template);
        } else {
            $this->excelObj = new PHPExcel();
        }
    }

    public function setFileType($type = 'xls') {
        if(!in_array($type, array('xls', 'xlsx'))) $type = 'xls';
        $this->fileType = $type;
        return $this;
    }

    public function setFileName($name = '') {
        if(empty($name)) $name = "default.{$this->fileType}";

        if(strpos($_SERVER['HTTP_USER_AGENT'], 'MSIE') === false) {
            $this->fileName = $name;
        } else {
            $this->fileName = urlencode($name);
        }

        return $this;
    }

    public function setTemplate($template = null) {
        //Todo 判断是否路径格式
        $this->template = (!$template) ? null : $template;
        return $this;
    }

    public function setIsLimitMemory($limit = false) {
        $this->isLimitMemory = ($limit != true) ? false : true;
        return $this;
    }

    public function setIsLimitTime($limit = false) {
        $this->isLimitTime = ($limit != true) ? false : true;
        return $this;
    }

    public function write(array $data = array()) {
        if($this->isLimitMemory) ini_set('memory_limit','2024M');
        if($this->isLimitTime) set_time_limit(0);

        $activeSheet = $this->excelObj->getActiveSheet(0);
        foreach($data as $cell => $val) {
            $activeSheet->setCellValue($cell, $val);
        }
    }

    /**$data =  array(
        array('sheet' => 0, 'name' => 'sheet name 0', 'data' => array(...),
            'startRow' => 4, 'appendRows' => 112, 
        ),
        array('sheet' => 1, 'name' => 'sheet name 1', 'data' => array(...), ...),
        ...
    );*/
    public function writeMultiSheet(array $data = array()) {
        if($this->isLimitMemory) ini_set('memory_limit','2024M');
        if($this->isLimitTime) set_time_limit(0);

        foreach($data as $value) {
            $index = $value['sheet'];
            $total_sheet = $this->excelObj->getSheetCount();
            if($index >= $total_sheet) { //sheet 不存在
                for($i = $total_sheet; $i <= $index; $i++) {
                    $this->excelObj->createSheet($index);
                }
            }
            $this->excelObj->setActiveSheetIndex($index);
            $activeSheet = $this->excelObj->getActiveSheet();
            if($value['name']) $activeSheet->setTitle($value['name']);
            foreach($value['data'] as $cell => $val) {
                $activeSheet->setCellValue($cell, $val);
            }
        }
    }

    // 导出excel
    public function output(array $data = array(), $simple = true) {
        $excelType = ($this->fileType === "xlsx") ? "Excel2007" : "Excel5";
        $objWriter = PHPExcel_IOFactory::createWriter($this->excelObj, $excelType);
        $objWriter->setPreCalculateFormulas(false);// 禁用公式预先计算

        if($simple) {
            $this->write($data);
        } else {
            $this->writeMultiSheet($data);
        }

        ob_end_clean();
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; charset=utf-8');
        header('Content-Disposition: attachment;filename="'.$this->fileName.'"');
        header('Cache-Control: max-age=0');
        $objWriter->save('php://output');
        exit;
    }

}

<?php
namespace Jingwu\Excel;

/**
 * Class Excel_XML
 * 用于将数组数据转储为Excel可读格式的简单导出库，支持OpenOffice Calc
 * @author    jingwu@vip.163.com
 */
class Excel_XML {

        /**
         * Excel的MicrosoftXML头文件
         * @var string
         */
        const sHeader = "<?xml version=\"1.0\" encoding=\"%s\"?\>\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">";

        /**
         * Excel的MicrosoftXML尾文件
         * @var string
         */
        const sFooter = "</Workbook>";

        /**
         * 工作表数据
         * @var array
         */
        private $aWorksheetData;

        /**
         * 输出字符串
         * @var string
         */
        private $sOutput;

        /**
         * 浏览器使用的编码
         * @var string
         */
        private $sEncoding;

        /**
         * 声明允许用户定义编码的类。
         * @param string $sEncoding 使用的编码
         */
        public function __construct($sEncoding = 'UTF-8') {
                $this->sEncoding = $sEncoding;
                $this->sOutput = '';
        }

        /**
         * 添加一个工作表
         * @param string $title 标题
         * @param array $data 2-dimensional array of data
         */
        public function addWorksheet($title, $data) {
                $this->aWorksheetData[] = array(
                        'title' => $this->getWorksheetTitle($title),
                        'data'  => $data
                );
        }

        /**
         * 发送到浏览器
         * @param string $filename header中的文件名
         */
        public function sendWorkbook($filename) {
                if (!preg_match('/\.(xml|xls)$/', $filename)):
                        throw new Exception('Filename mimetype must be .xml or .xls');
                endif;
                $filename = $this->getWorkbookTitle($filename);
                $this->generateWorkbook();
                if (preg_match('/\.xls$/', $filename)):
                        header("Content-Type: application/vnd.ms-excel; charset=" . $this->sEncoding);
                        header("Content-Disposition: inline; filename=\"" . $filename . "\"");
                else:
                        header("Content-Type: application/xml; charset=" . $this->sEncoding);
                        header("Content-Disposition: attachment; filename=\"" . $filename . "\"");
                endif;
                echo $this->sOutput;
        }

        /**
         * 写入文件, 请确保文件可写且不存在
         * @param string $filename 文件名(含后缀)
         * @param string $path 路径，可选
         */
        public function writeWorkbook($filename, $path = '') {
                $this->generateWorkbook();
                $filename = $this->getWorkbookTitle($filename);
                if (!$handle = @fopen($path . $filename, 'w+')):
                        throw new Exception(sprintf("Not allowed to write to file %s", $path . $filename));
                endif;
                if (@fwrite($handle, $this->sOutput) === false):
                        throw new Exception(sprintf("Error writing to file %s", $path . $filename));
                endif;
                @fclose($handle);
                return sprintf("File %s written", $path . $filename);
        }

        /**
         * 取得工作表, 字符串
         * @return string 输出生成的xml字符串
         */
        public function getWorkbook() {
                $this->generateWorkbook();
                return $this->sOutput;
        }

        /**
         * 修正工作簿标题（如必要）去掉非法字符。
         * @param string $filename 文件名
         * @return string Corrected filename
         */
        private function getWorkbookTitle($filename) {
                return preg_replace('/[^aA-zZ0-9\_\-\.]/', '', $filename);
        }

        /**
         * 修正工作表标题（用户给定）
         * @param string $title Desired worksheet title
         * @return string Corrected worksheet title
         */
        private function getWorksheetTitle($title) {
                $title = preg_replace ("/[\\\|:|\/|\?|\*|\[|\]]/", "", $title);
                return substr ($title, 0, 31);
        }

        /**
         * 生成工作簿
         */
        private function generateWorkbook() {
                $this->sOutput .= stripslashes(sprintf(self::sHeader, $this->sEncoding)) . "\n";
                foreach ($this->aWorksheetData as $item):
                        $this->generateWorksheet($item);
                endforeach;
                $this->sOutput .= self::sFooter;
        }

        /**
         * 生成工作表。超过Excel的最大行数时，进行切片。
         * @param array $item 工作表数据
         * @todo 检查是否为数组
         */
        private function generateWorksheet($item) {
                $this->sOutput .= sprintf("<Worksheet ss:Name=\"%s\">\n    <Table>\n", $item['title']);
                if (count($item['data']))
                        $item['data'] = array_slice($item['data'], 0, 65536);
                foreach ($item['data'] as $k => $v):
                        $this->generateRow($v);
                endforeach;
                $this->sOutput .= "    </Table>\n</Worksheet>\n";
        }

        /**
         * 生成一行
         * @param array $iem 一行数据
         */
        private function generateRow($item) {
                $this->sOutput .= "        <Row>\n";
                foreach ($item as $k => $v):
                        $this->generateCell($v);
                endforeach;
                $this->sOutput .= "        </Row>\n";
        }

        /**
         * 生成单元格
         * @param string $item 单元格数据
         */
        private function generateCell($item) {
                $type = 'String';
                if (is_numeric($item)):
                        $type = 'Number';
                        if ($item{0} == '0' && strlen($item) > 1 && $item{1} != '.'):
                                $type = 'String';
                        endif;
                endif;
                $item = str_replace('&#039;', '&apos;', htmlspecialchars($item, ENT_QUOTES));
                $this->sOutput .= sprintf("            <Cell><Data ss:Type=\"%s\">%s</Data></Cell>\n", $type, $item);
        }

        /**
         * 析构重置变量
         */
        public function __destruct() {
                unset($this->aWorksheetData);
                unset($this->sOutput);
        }

}

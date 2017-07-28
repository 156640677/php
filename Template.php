<?php

/**
 * PHPWord
 *
 * Copyright (c) 2010 PHPWord
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPWord
 * @package    PHPWord
 * @copyright  Copyright (c) 010 PHPWord
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    Beta 0.6.2, 24.07.2010
 */

/**
 * PHPWord_DocumentProperties
 *
 * @category   PHPWord
 * @package    PHPWord
 * @copyright  Copyright (c) 2009 - 2010 PHPWord (http://www.codeplex.com/PHPWord)
 */
class PHPWord_Template {
    private $_objZip;
    private $_tempFileName;
    private $_documentXML;
    private $_header1XML;
    private $_footer1XML;
    private $_rels;
    private $_types;
    private $_countRels;

    //private $_domXML;
    /**
     * Document in tables
     * @var DOMNode
     */
    private $_tables;

    private $_templateRow;

    private static $doc_tag_key = 'w:';

    private static $doc_tag_key_temp = 'w_';
    /**
     * Create a new Template Object
     * @param string $strFilename
     */
    public function __construct($strFilename) {
        $path = dirname($strFilename);

        $this->_tempFileName = $path . DIRECTORY_SEPARATOR . time() . '.docx'; // $path doesn't include the trailing slash - Custom code by Matt Bowden (blenderstyle) 04/12/2011

        copy($strFilename, $this->_tempFileName); // Copy the source File to the temp File

        $this->_objZip = new ZipArchive();
        $this->_objZip->open($this->_tempFileName);

        $this->_documentXML = $this->_objZip->getFromName('word/document.xml');
        $this->_header1XML  = $this->_objZip->getFromName('word/header1.xml'); // Custom code by Matt Bowden (blenderstyle) 04/12/2011
        $this->_footer1XML  = $this->_objZip->getFromName('word/footer1.xml'); // Custom code by Matt Bowden (blenderstyle) 04/12/2011
        $this->_rels        = $this->_objZip->getFromName('word/_rels/document.xml.rels'); #erap 07/07/2015
        $this->_types       = $this->_objZip->getFromName('[Content_Types].xml'); #erap 07/07/2015
        $this->_countRels   = substr_count($this->_rels, 'Relationship') - 1; #erap 07/07/2015
    }

    /**
     * Set a Template value
     * @param mixed $search
     * @param mixed $replace
     */
    public function setValue($search, $replace) {
        $search = $this->getSignKey($search);
        $replace = $this->limpiarString($replace);
        $this->_documentXML = str_replace($search, $replace, $this->_documentXML);
        $this->_header1XML = str_replace($search, $replace, $this->_header1XML); // Custom code by Matt Bowden (blenderstyle) 04/12/2011
        $this->_footer1XML = str_replace($search, $replace, $this->_footer1XML); // Custom code by Matt Bowden (blenderstyle) 04/12/2011
    }

    /**
     * Save Template
     * @param string $strFilename
     * @throws Exception
     */
    public function save($strFilename) {
        if (file_exists($strFilename))
            unlink($strFilename);
        //clear template sign
        $this->clearTemplateSign();

        $this->_objZip->addFromString('word/document.xml', $this->_documentXML);
        $this->_objZip->addFromString('word/header1.xml', $this->_header1XML); // Custom code by Matt Bowden (blenderstyle) 04/12/2011
        $this->_objZip->addFromString('word/footer1.xml', $this->_footer1XML); // Custom code by Matt Bowden (blenderstyle) 04/12/2011
        $this->_objZip->addFromString('word/_rels/document.xml.rels', $this->_rels); #erap 07/07/2015
        $this->_objZip->addFromString('[Content_Types].xml', $this->_types); #erap 07/07/2015
        // Close zip file
        if ($this->_objZip->close() === false)
            throw new Exception('Could not close zip file.');

        rename($this->_tempFileName, $strFilename);
    }

    public function replaceImage($path, $imageName) {
        $this->_objZip->deleteName('word/media/' . $imageName);
        $this->_objZip->addFile($path, 'word/media/' . $imageName);
    }

    /**
     * $a002 = array(array('img' => 'image/002.jpg','size' => array(50, 50)));
     * $document->replaceStrToImg( 'image002', $a002);
     * @param $strKey
     * @param $arrImgPath
     * @param $isPHPWord
     */
    public function replaceStrToImg( $strKey, $arrImgPath,$isPHPWord=0){
        //289x108
        $strKey = $this->getSignKey($strKey);
        $relationTmpl = '<Relationship Id="RID" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/IMG"/>';
        $imgTmpl = '<w:pict><v:shape type="#_x0000_t75" style="width:WIDpx;height:HEIpx"><v:imagedata r:id="RID" o:title=""/></v:shape></w:pict>';
        $typeTmpl = ' <Override PartName="/word/media/IMG" ContentType="image/EXT"/>';
        $toAdd = $toAddImg = $toAddType = '';
        $aSearch = array('RID', 'IMG');
        $aSearchType = array('IMG', 'EXT');
        foreach($arrImgPath as $img){
            $tt = explode('.', $img['img']);
            $imgExt = array_pop($tt);
            if( in_array($imgExt, array('jpg', 'JPG') ) )
                $imgExt = 'jpeg';
            $imgName = 'img' . $this->_countRels . '.' . $imgExt;
            $rid = 'rId' . $this->_countRels++;

            $this->_objZip->addFile($img['img'], 'word/media/' . $imgName);
            if( isset($img['size']) ){
                $w = $img['size'][0];
                $h = $img['size'][1];
            }
            else{
                $w = 289;
                $h = 108;
            }

            $toAddImg .= str_replace(array('RID', 'WID', 'HEI'), array($rid, $w, $h), $imgTmpl) ;
            if( isset($img['dataImg']) )
                $toAddImg .= '<w:br/><w:t>' . $this->limpiarString($img['dataImg']) . '</w:t><w:br/>';

            $aReplace = array($imgName, $imgExt);
            $toAddType .= str_replace($aSearchType, $aReplace, $typeTmpl) ;

            $aReplace = array($rid, $imgName);
            $toAdd .= str_replace($aSearch, $aReplace, $relationTmpl);
        }
        /*
         //如果直接用代码生成文件替换的则采用如下方法
         if($isPHPWord){
            //用PHPWord的addText方法生成的内容多了一个属性：xml:space="preserve"
            //详情请见PHPWord源码中Base.php的69行
            $imgSign = '<w:t xml:space="preserve">' . $strKey .'</w:t>';
        }*/
        //如果用代码生成文件后再编辑，再替换，则采用如下方法
        $imgSign = '<w:t xml:space="preserve">' . $strKey .'</w:t>';
        $pos = strpos($this->_documentXML, $imgSign);
        if($pos === false){
            $imgSign = '<w:t>' . $strKey .'</w:t>';
        }
        $this->_documentXML = str_replace($imgSign, $toAddImg, $this->_documentXML);
        $this->_types       = str_replace('</Types>', $toAddType, $this->_types) . '</Types>';
        $this->_rels        = str_replace('</Relationships>', $toAdd, $this->_rels) . '</Relationships>';
    }

    function limpiarString($str) {
        return str_replace(
            array('&', '<', '>', "\n"),
            array('&amp;', '&lt;', '&gt;', "\n" . '<w:br/>'),
            $str
        );
    }

    public function setImg($data){
        foreach ($data as $key=>$value){
            $this->replaceStrToImg('img.'.$key,array($value));
        }
    }
    /**
     * 获取所有Table
     */
    private function getTables(){
        $tables = null;
        if(!$this->_tables){
            $doc = new DOMDocument();
            $xml = str_replace(PHPWord_Template::$doc_tag_key,PHPWord_Template::$doc_tag_key_temp, $this->_documentXML);
            $doc->loadXML($xml);
            $nodeList = $doc->getElementsByTagName(PHPWord_Template::$doc_tag_key_temp."tbl");
            foreach ($nodeList as $node) {
                $this->_tables[] = $node;
            }
        }
        return $this->_tables;
    }

    /**
     * 获取指定Table
     */
    public function getTable($sign){
        $k = $this->getSignKey($sign);
        $tables = $this->getTables();
        foreach ($tables as $key=>$value){
            $isTable = $this->isTable($value,$k);
            if($isTable){
                return $value;
            }
        }
        return null;
    }

    /**
     * 是否为当前table
     * @param $table
     * @param $sign
     * @return bool
     */
    public function isTable($table,$sign){
        $str = $this->getTableTemplateRow($table,$sign);
        $isTable = strpos($str,$sign);
        if ($isTable==false)return false;
        $this->_templateRow[$sign]=$str;
        return true;
    }

    private function getTableTemplateRow($table,$sign){
        $templateRow = isset($this->_templateRow[$sign]);
        if($templateRow)return $this->_templateRow[$sign];
        $node = $table->lastChild;
        $str = $node->ownerDocument->saveXML( $node );
        return $str;
    }

    public function addRow($table,$sign,$data){
        $sign = $this->getSignKey($sign);
        $row = $this->getTableTemplateRow($table,$sign);
        $key = substr($sign,0,strrpos($sign,"."));
        $b_str = $row;
        foreach ($data as $k=>$v){
            $tkey =$key.'.'.$k."}";//${table2.prop}
            $b_str = str_replace($tkey,$v,$b_str);
            //var_dump($b_str);
        }
        $this->appendXML($row,$b_str);
    }

    private function appendXML($templateRow,$row){
        $newTmplateRow = str_replace(PHPWord_Template::$doc_tag_key_temp,PHPWord_Template::$doc_tag_key, $templateRow);
        $newRow = str_replace(PHPWord_Template::$doc_tag_key_temp,PHPWord_Template::$doc_tag_key, $row);
        //保留模板行
        $this->_documentXML = str_replace($newTmplateRow."</w:tbl>",$newRow.$newTmplateRow."</w:tbl>",$this->_documentXML);
    }

    private function clearTemplateSign(){
        $this->cleanTemplateRow();
        $this->cleanVar();
    }
    private function cleanTemplateRow(){
        if(count($this->_templateRow)==0)return;
        foreach ($this->_templateRow as $key=>$value){
            $newTemplateRow = str_replace(PHPWord_Template::$doc_tag_key_temp,PHPWord_Template::$doc_tag_key, $value);
            //var_dump($newTemplateRow);
            $this->_documentXML = str_replace($newTemplateRow."</w:tbl>","</w:tbl>", $this->_documentXML);
        }
    }

    private function cleanVar(){
        //$preg = '/\$\{([a-z]|[0-9])*+\.+([a-z]|[0-9])*+\}/';
        $preg = '/(?:\$\{+\S+\})/';
        //$newStr = preg_replace($preg,"=",$str);
        //$index = strrpos($this->_documentXML,'${');
        //$s = substr($this->_documentXML,$index,$index+200);
        //var_dump($s);
        $this->_documentXML = preg_replace($preg,"", $this->_documentXML);
    }

    private function getSignKey($sign){
        return '${'.$sign."}";
    }
}

?>
<?php
include_once('tbszip.php');


class FrameWord{
    public $archive;
    public $newArchive;
    private $text;
    
    public function OpenFile(){
        $zip = new clsTbsZip();
        $zip->Open($this->archive);
        $this->text = $zip->FileRead('word/document.xml');
    }
    
    public function InsertWord($word, $newWord){
        $p = strpos($this->text, $word);
        $pI = strpos(strrev($this->text), strrev($word));
        $this->text = substr_replace($this->text, '', ($p), -$pI);
        $this->text = substr_replace($this->text, $newWord, ($p), -($pI+ strlen($newWord)));
    }
    
    public function GetText(){
        return $this->text;
    }
    
    public function SaveDocx(){
        $zip = new clsTbsZip();
        $zip->Open($this->archive);
        $zip->FileReplace('word/document.xml', $this->text, TBSZIP_STRING);
        $zip->Flush(TBSZIP_FILE, $this->newArchive);
    }
    
}
?>

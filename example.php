<?php
include_once('FrameWord/frameword.php');

$frameword = new FrameWord();
$frameword->archive = "test.docx";
$frameword->newArchive = "novo.docx";
$frameword->OpenFile();
echo $frameword->GetText()."<br/><br/><br/>";

$frameword->InsertWord("<w:t>&lt;profissão&gt;</w:t>", "<w:t>Programador</w:t>");
echo $frameword->GetText()."<br/><br/><br/>";

$frameword->SaveDocx();
?>

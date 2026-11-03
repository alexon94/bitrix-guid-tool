<?php
namespace GuidTool;
final class Excel {
    public static function col($n) { $s=''; for ($n++;$n>0;$n=intdiv($n-1,26)) $s=chr(65+(($n-1)%26)).$s; return $s; }
    public static function index($s) { $n=0; foreach(str_split($s) as $c) $n=$n*26+ord($c)-64; return $n-1; }
    private static function xml($s) { if (stripos($s,'<!DOCTYPE')!==false || stripos($s,'<!ENTITY')!==false) throw new \RuntimeException('Запрещён DTD в XLSX'); $old=libxml_use_internal_errors(true); try {$x=simplexml_load_string($s,'SimpleXMLElement',LIBXML_NONET|LIBXML_NOCDATA); if (!$x) throw new \RuntimeException('Некорректный XML XLSX'); return $x;} finally {libxml_clear_errors();libxml_use_internal_errors($old);} }
    private static function part($z,$p) { $s=$z->statName($p); if (!$s || $s['size']>50000000) throw new \RuntimeException('Отсутствует/слишком большой раздел XLSX: '.$p); return self::xml($z->getFromName($p)); }
    public static function dateStyles($z) {
        if($z->locateName('xl/styles.xml')===false)return [];
        $styles=self::part($z,'xl/styles.xml');$custom=[];
        foreach($styles->xpath('//*[local-name()="numFmts"]/*[local-name()="numFmt"]') as $fmt)$custom[(int)$fmt['numFmtId']]=(string)$fmt['formatCode'];
        $builtins=array_merge(range(14,22),range(27,36),range(50,58));$dates=[];
        $xfs=$styles->xpath('//*[local-name()="cellXfs"]/*[local-name()="xf"]');
        foreach($xfs as $idx=>$xf){$id=(int)$xf['numFmtId'];if(in_array($id,$builtins,true)){$dates[$idx]=true;continue;}
            if(!isset($custom[$id]))continue;
            $fmt=$custom[$id];$fmt=preg_replace('/"(?:[^"]|"")*"/u','',$fmt);
            $fmt=preg_replace('/\\./u','',$fmt);
            $fmt=preg_replace('/\[[^\]]*\]/u','',$fmt);
            if(preg_match('/[dy]/i',$fmt)||preg_match('/(?:m{3,}|m{1,2}[^a-z]*[dyy]|[dyy][^a-z]*m{1,2})/i',$fmt))$dates[$idx]=true;
        }
        return $dates;
    }
    public static function dateValue($serial,$date1904=false) {
        if(!is_numeric($serial)||!is_finite((float)$serial))return '';
        $days=(int)floor((float)$serial);
        if($days<0||$days>2958465)return '';
        $base=$date1904?'1904-01-01':'1899-12-30';
        if(!$date1904&&$days<60)$days++;
        try{$dt=new \DateTimeImmutable($base,new \DateTimeZone('UTC'));return $dt->modify('+'.$days.' days')->format('d.m.Y');}catch(\Throwable $e){return '';}
    }
    public static function read($path) {
        if (!class_exists('ZipArchive')) throw new \RuntimeException('Требуется PHP zip'); $z=new \ZipArchive(); if($z->open($path)!==true) throw new \RuntimeException('XLSX не открывается');
        try {
            $wb=self::part($z,'xl/workbook.xml'); $sheets=$wb->xpath('//*[local-name()="sheet"]'); if (!$sheets) throw new \RuntimeException('Нет листов');
            $rid=(string)$sheets[0]->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships')->id;
            $rels=self::part($z,'xl/_rels/workbook.xml.rels'); $target=''; foreach($rels->children() as $rel) if((string)$rel['Id']===$rid && (string)$rel['TargetMode']!=='External') $target=(string)$rel['Target'];
            $p=($target!=='' && $target[0]==='/')?ltrim($target,'/'): 'xl/'.$target; $p=preg_replace('~^xl/xl/~','xl/',$p);
            if (!$target || strpos($p,'..')!==false || strpos($p,'xl/')!==0) throw new \RuntimeException('Неподдерживаемый путь листа');
            $ss=[]; if ($z->locateName('xl/sharedStrings.xml')!==false) foreach(self::part($z,'xl/sharedStrings.xml')->xpath('//*[local-name()="si"]') as $si) {$v='';foreach($si->xpath('.//*[local-name()="t"]') as $t) $v.=(string)$t;$ss[]=$v;}
            $dateStyles=self::dateStyles($z);$date1904=(string)($wb->xpath('//*[local-name()="workbookPr"]')[0]['date1904']??'')==='1';
            $rows=[]; $max=-1;
            foreach(self::part($z,$p)->xpath('//*[local-name()="sheetData"]/*[local-name()="row"]') as $r) {
                $n=(int)$r['r']; $cells=[];
                foreach($r->children() as $c) {
                    if($c->getName()!=='c')continue; if(!preg_match('/^([A-Z]+)[0-9]+$/',(string)$c['r'],$m))continue; $i=self::index($m[1]); if($i>255)throw new \RuntimeException('Максимум 256 колонок');$max=max($max,$i);
                    $t=(string)$c['t']; $v=$c->xpath('./*[local-name()="v"]');$v=$v?(string)$v[0]:'';$error='';
                    if($c->xpath('./*[local-name()="f"]'))$error='Формула: замените значением';
                    if($t==='s')$v=$ss[(int)$v]??''; elseif($t==='inlineStr'){$v='';foreach($c->xpath('.//*[local-name()="t"]') as $x)$v.=(string)$x;}
                    if($t==='e')$error='Ошибка Excel'; $cells[$i]=['v'=>$v,'type'=>$t,'error'=>$error,'numeric'=>($t===''||$t==='n')&&$v!=='','date'=>($t==='d'||(($t===''||$t==='n')&&isset($dateStyles[(int)$c['s']]))),'date_value'=>$t==='d'?(preg_match('/^\d{4}-\d{2}-\d{2}/',$v)?substr($v,8,2).'.'.substr($v,5,2).'.'.substr($v,0,4):$v):((($t===''||$t==='n')&&isset($dateStyles[(int)$c['s']]))?self::dateValue($v,$date1904):'')];
                }
                if($cells)$rows[$n]=$cells; if(count($rows)>20000)throw new \RuntimeException('Максимум 20 000 строк');
            }
            ksort($rows);return ['sheet'=>(string)$sheets[0]['name'],'rows'=>$rows,'width'=>$max+1];
        }finally{$z->close();}
    }
    private static function esc($v) {return htmlspecialchars((string)$v,ENT_QUOTES|ENT_XML1,'UTF-8');}
    public static function write($path,$headers,$rows,$statusCol=null) {
        $z=new \ZipArchive();if($z->open($path,\ZipArchive::CREATE|\ZipArchive::OVERWRITE)!==true)throw new \RuntimeException('Не удалось создать XLSX');
        $z->addFromString('[Content_Types].xml','<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
        $z->addFromString('_rels/.rels','<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
        $z->addFromString('xl/workbook.xml','<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Отчёт" sheetId="1" r:id="rId1"/></sheets></workbook>');
        $z->addFromString('xl/_rels/workbook.xml.rels','<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
        $z->addFromString('xl/styles.xml','<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/></font></fonts><fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDBEAFE"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFE8A3"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFECACA"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="4" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/></cellXfs></styleSheet>');
        $f=fopen('php://temp/maxmemory:4194304','w+');fwrite($f,'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>');
        $emit=function($values,$n,$head=false)use($f,$statusCol){fwrite($f,'<row r="'.$n.'">');foreach($values as $i=>$v){$s=0;if($head)$s=$statusCol!==null&&$i>=$statusCol?4:1;elseif($statusCol!==null&&$i>=$statusCol){$vstr=(string)($values[$statusCol]??'');$s=(strpos($vstr,'Обновлено')!==false||strpos($vstr,'Успешно')!==false)?2:((strpos($vstr,'Ошибка')!==false||strpos($vstr,'частичная')!==false)?5:3);}fwrite($f,'<c r="'.Excel::col($i).$n.'" s="'.$s.'" t="inlineStr"><is><t xml:space="preserve">'.Excel::esc($v).'</t></is></c>');}fwrite($f,'</row>');};
        $emit(array_values($headers),1,true);$n=2;foreach($rows as $row)$emit(array_values($row),$n++);fwrite($f,'</sheetData></worksheet>');rewind($f);$z->addFromString('xl/worksheets/sheet1.xml',stream_get_contents($f));fclose($f);$z->close();
    }
}

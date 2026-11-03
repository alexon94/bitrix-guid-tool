<?php
namespace GuidTool;
final class CoreStats {
    public static function count($job){$a=['total'=>count($job['rows']??[]),'ready'=>0,'will_skip'=>0,'success'=>0,'errors'=>0,'skipped'=>0,'remaining'=>0];foreach($job['rows']??[] as $r){$eligible=in_array($r['status']??'',['Готов к обновлению','GUID уже совпадает'],true)&&empty($r['date_issue']);if($eligible)$a['ready']++;else $a['will_skip']++;$result=$r['result']??'';if($result==='К обновлению'||$result==='Не обработано')$a['remaining']++;elseif($result==='Обновлено')$a['success']++;elseif(strpos($result,'Ошибка')===0||strpos($result,'Запись начата')===0)$a['errors']++;elseif($result==='Пропустить'||strpos($result,'Пропущено')===0)$a['skipped']++;}return $a;}
}

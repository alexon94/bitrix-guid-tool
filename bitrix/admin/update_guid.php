<?php
/** Bitrix GUID Tool 1.1.8 — admin entry. */
require_once $_SERVER['DOCUMENT_ROOT'].'/bitrix/modules/main/include/prolog_admin_before.php';
if(!$USER->IsAdmin())$APPLICATION->AuthForm('Только для администраторов');
if(!\Bitrix\Main\Loader::includeModule('iblock')||!\Bitrix\Main\Loader::includeModule('catalog'))die('Нужны iblock и catalog');
if(defined('LANG_CHARSET')&&strcasecmp(LANG_CHARSET,'UTF-8')!==0)die('Требуется UTF-8');
foreach(['Excel','Core','CoreStats','Store'] as $lib)require_once $_SERVER['DOCUMENT_ROOT'].'/local/bitrix-guid-tool/lib/'.$lib.'.php';
use GuidTool\Core as C;use GuidTool\Excel as X;use GuidTool\CoreStats as S;use GuidTool\Store;
function h($v){return htmlspecialchars((string)$v,ENT_QUOTES|ENT_SUBSTITUTE,'UTF-8');}
function form($action,$profile){echo '<form method="post" enctype="multipart/form-data">'.bitrix_sessid_post().'<input type="hidden" name="profile" value="'.h($profile).'"><input type="hidden" name="action" value="'.h($action).'"><input type="hidden" name="tab_token" value="'.h($_REQUEST['tab_token']??'').'">';}
function getval($r,$i){return (string)($r['cells'][$i]['v']??'');}
function reportFile($job,$type){
    $width=$job['data']['width'];$headers=[];for($i=0;$i<$width;$i++)$headers[]=(string)($job['data']['rows'][1][$i]['v']??X::col($i));
    $extra=['ID Битрикс','Тип','Название на сайте','XML_ID на сайте','Базовая цена сайта','Статус сверки','Причины','Результат обновления'];$rows=[];
    foreach($job['rows'] as $r){if($r['exclude'])continue;$vals=[];for($i=0;$i<$width;$i++)$vals[]=($type==='prepared'&&$i===$job['map']['article']&&!empty($job['opts']['normalize']))?$r['article']:getval($r,$i);
        if($type!=='prepared')$vals=array_merge($vals,[$r['id'],$r['type'],$r['name'],$r['old_guid'],$r['site_price'],$r['status'],$r['issues'],$type==='preliminary'?($r['initial_result']??($r['result']==='К обновлению'?'К обновлению':'Пропустить')):$r['result']]);$rows[]=$vals;}
    $tmp=tempnam(sys_get_temp_dir(),'guid-xlsx-');if(!$tmp)throw new RuntimeException('Не удалось создать временный файл');try{X::write($tmp,$type==='prepared'?$headers:array_merge($headers,$extra),$rows,$type==='prepared'?null:$width);}catch(Throwable $e){@unlink($tmp);throw $e;}return $tmp;
}
function cacheDuplicates($store,$profile,$job){
    $dups=\GuidTool\Core::duplicates($job['rows']);
    $width=$job['data']['width'];$headers=[];
    for($i=0;$i<$width;$i++)$headers[]=(string)($job['data']['rows'][1][$i]['v']??\GuidTool\Excel::col($i));
    $rows=[];
    foreach($dups as $r){
        $values=[];
        for($i=0;$i<$width;$i++)$values[]=getval($r,$i);
        $rows[]=$values;
    }
    $tmp=tempnam(sys_get_temp_dir(),'guid-dups-');if(!$tmp)throw new RuntimeException('Не удалось создать файл дублей');
    try{\GuidTool\Excel::write($tmp,$headers,$rows);$store->duplicatesSave($profile,$tmp);}finally{@unlink($tmp);}
}
function sendXlsx($path,$name,$delete=false){global $APPLICATION;$APPLICATION->RestartBuffer();header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');header('Content-Disposition: attachment; filename="'.$name.'"');header('Content-Length: '.filesize($path));header('Cache-Control: no-store');readfile($path);if($delete)@unlink($path);exit;}
function archiveCurrent($store,$profile,$job){$tmp=reportFile($job,'final');try{$store->archiveSave($profile,$job,$tmp);}finally{@unlink($tmp);}}
$profile=(string)(($_SERVER['REQUEST_METHOD']??'GET')==='POST'?($_POST['profile']??''):($_GET['profile']??''));$isAjax=(($_POST['ajax']??'')==='1'&&($_POST['action']??'')==='batch');$tabToken=(string)($_REQUEST['tab_token']??'');$error='';$message='';$job=[];$cfg=null;$store=null;
try{
    if($profile!==''){$cfg=C::config($profile);$store=new Store($_SERVER['DOCUMENT_ROOT']);$job=$store->load($profile);$link=C::validate($cfg);if(($job['tab_token']??'')!==$tabToken||$tabToken==='')$job=[];}
    if(($_SERVER['REQUEST_METHOD']??'GET')==='POST'){
        if(!$cfg)throw new RuntimeException('Сначала выберите сайт');if(!check_bitrix_sessid())throw new RuntimeException('Истекла сессия');$action=$_POST['action']??'';
        if($action==='upload'){
            $f=$_FILES['xlsx']??[];if(($f['error']??1)!==UPLOAD_ERR_OK||($f['size']??0)>10485760||strtolower(pathinfo($f['name']??'',PATHINFO_EXTENSION))!=='xlsx'||!is_uploaded_file($f['tmp_name']))throw new RuntimeException('Загрузите XLSX до 10 МБ');
            if(!preg_match('/^[a-f0-9]{32}$/D',$tabToken))throw new RuntimeException('Не определена рабочая вкладка');
            $lock=$store->lock($profile);try{$previous=$store->load($profile);if(!empty($previous['archive_id'])&&in_array($previous['phase']??'',['processing','finished'],true))archiveCurrent($store,$profile,$previous);$store->duplicatesClear($profile);$job=['tab_token'=>$tabToken,'phase'=>'uploaded','data'=>X::read($f['tmp_name']),'created'=>date('c'),'profile'=>$profile,'token'=>bin2hex(random_bytes(16))];$store->save($profile,$job);}finally{$store->unlock($lock);}
        }
        elseif(in_array($action,['prepare_preview','prepare_apply','check_direct','check_dates'],true)){
            if(!in_array($job['phase']??'',['uploaded','preview','prepared'],true))throw new RuntimeException('Сначала загрузите XLSX');$map=[];
            foreach(['article','guid','name','price'] as $k){$map[$k]=(int)($_POST[$k]??-1);if($map[$k]<0||$map[$k]>=$job['data']['width'])throw new RuntimeException('Неверная колонка '.$k);}if(count(array_unique($map))!==4)throw new RuntimeException('Колонки должны различаться');
            $map['start']=max(2,(int)($_POST['start']??2));$opts=['drop_price'=>!empty($_POST['drop_price']),'normalize'=>!empty($_POST['normalize'])];
            if($action==='check_direct')$opts=['drop_price'=>false,'normalize'=>false];
            if($action==='check_dates'){$found=[];foreach($job['data']['rows'] as $n=>$cells){if($n<$map['start'])continue;$issue=C::dateIssue($cells[$map['article']]??[]);if($issue!=='')$found[]=['n'=>$n,'issue'=>$issue];}$job['date_check']=$found;$job['map']=$map;$job['opts']=$opts;$message='Проверка артикулов Excel завершена: найдено '.count($found).' ячеек в формате даты.';$store->save($profile,$job);}else{unset($job['date_check']);[$rows,$stats]=C::prepare($job['data'],$map,$opts);$job['map']=$map;$job['opts']=$opts;$job['stats']=$stats;$job['rows']=$rows;$job['prepared_applied']=($action==='prepare_apply');cacheDuplicates($store,$profile,$job);
            if($action==='prepare_preview')$job['phase']='preview';elseif($action==='prepare_apply')$job['phase']='prepared';else{C::plan($job,$cfg,$link);$job['rows']=array_values(array_filter($job['rows'],function($r){return empty($r['exclude']);}));foreach($job['rows'] as &$r)$r['initial_result']=$r['result'];unset($r);}
            $store->save($profile,$job);}
        }
        elseif($action==='plan'){$lock=$store->lock($profile);try{$job=$store->load($profile);if(($job['phase']??'')!=='prepared')throw new RuntimeException('Сначала примените подготовку');C::plan($job,$cfg,$link);$job['rows']=array_values(array_filter($job['rows'],function($r){return empty($r['exclude']);}));foreach($job['rows'] as &$r)$r['initial_result']=$r['result'];unset($r);$store->save($profile,$job);}finally{$store->unlock($lock);}}
        elseif($action==='batch'){
            $lock=$store->lock($profile);try{
                $job=$store->load($profile);if(($job['tab_token']??'')!==$tabToken||$tabToken==='')throw new RuntimeException('Это задание принадлежит другой вкладке');if($isAjax&&(!isset($_POST['expected_done'])||(int)$_POST['expected_done']!==S::count($job)['success']+S::count($job)['errors']))throw new RuntimeException('Состояние изменилось. Обновите страницу и проверьте статистику перед следующей партией');if(!in_array($job['phase']??'',['planned','processing'],true))throw new RuntimeException('Нет подготовленного плана или обновление уже завершено');
                if(empty($job['archive_id'])){$job['archive_id']=bin2hex(random_bytes(16));$job['started']=date('c');$job['admin']=(int)$USER->GetID();}
                $job['phase']='processing';$store->save($profile,$job);$processed=0;
                foreach($job['rows'] as &$r){if($r['result']!=='К обновлению')continue;if(!empty($r['date_issue'])){$r['result']='Пропущено: артикул сохранён как дата';continue;}if($processed>=25)break;$r['result']='Запись начата — проверить';$store->save($profile,$job);try{C::update($r,$job,$cfg,$link);}catch(Throwable $e){$r['result']='Ошибка/частичная запись: '.$e->getMessage();}$store->save($profile,$job);$processed++;}unset($r);
                if(S::count($job)['remaining']===0){$job['phase']='finished';$job['completed']=date('c');}
                $store->save($profile,$job);archiveCurrent($store,$profile,$job);$message='Обработан пакет: '.$processed.' строк.';
            }finally{$store->unlock($lock);}
        }
        elseif($action==='reset'){$lock=$store->lock($profile);try{$job=$store->load($profile);if(in_array($job['phase']??'',['processing','finished','stopped'],true)&&!empty($job['archive_id']))archiveCurrent($store,$profile,$job);$job=[];$store->save($profile,$job);$store->duplicatesClear($profile);}finally{$store->unlock($lock);}}
        elseif(in_array($action,['prepared_xlsx','preliminary_xlsx','final_xlsx'],true)){
            if(!isset($job['rows'],$job['data']))throw new RuntimeException('Нет данных для отчёта');$phase=$job['phase']??'';
            if($action==='prepared_xlsx'&&empty($job['prepared_applied']))throw new RuntimeException('Подготовка не выполнялась');
            if($action==='prepared_xlsx'&&!in_array($phase,['prepared','planned','processing','stopped','finished'],true))throw new RuntimeException('Подготовка не завершена');
            if($action==='preliminary_xlsx'&&!in_array($phase,['planned','processing','stopped','finished'],true))throw new RuntimeException('Нет сверки');
            if($action==='final_xlsx'&&!in_array($phase,['processing','stopped','finished'],true))throw new RuntimeException('Обновление ещё не начато');
            sendXlsx(reportFile($job,str_replace('_xlsx','',$action)),$action.'.xlsx',true);
        }
        elseif($action==='duplicates_xlsx'){
            if(!isset($job['rows'],$job['map']))throw new RuntimeException('Сначала проверьте подготовку');
            if(!C::duplicates($job['rows']))throw new RuntimeException('Дубли не найдены');
            sendXlsx($store->duplicatesFile($profile),'duplicates.xlsx');
        }
        elseif($action==='archive_download')sendXlsx($store->archivePath($profile,$_POST['id']??''),'guid-report-'.($_POST['id']??'').'.xlsx');
        elseif($action==='archive_delete'){
            if(($_POST['confirm']??'')!=='yes')throw new RuntimeException('Подтвердите удаление отчётов');$ids=$_POST['ids']??[];if(!is_array($ids)||!$ids)throw new RuntimeException('Выберите отчёты');
            $lock=$store->lock($profile);try{$store->archiveDelete($profile,$ids);}finally{$store->unlock($lock);}$message='Выбранные архивные отчёты удалены';
        }
    }
}catch(Throwable $e){$error=$e->getMessage();}
if($isAjax){$APPLICATION->RestartBuffer();header('Content-Type: application/json; charset=UTF-8');header('Cache-Control: no-store');if($error)http_response_code(409);$stats=$job?S::count($job):null;echo json_encode(['ok'=>$error==='', 'error'=>$error,'message'=>$message,'stats'=>$stats,'phase'=>$job['phase']??'', 'archive_ready'=>!empty($job['archive_id']),'processed'=>$processed??0],JSON_UNESCAPED_UNICODE|JSON_THROW_ON_ERROR);exit;}
$APPLICATION->SetTitle('Bitrix GUID Tool 1.1.8');require $_SERVER['DOCUMENT_ROOT'].'/bitrix/modules/main/include/prolog_admin_after.php';
if($error)\CAdminMessage::ShowMessage(h($error));if($message)echo '<p>'.h($message).'</p>';
// Описания согласованных пользовательских статусов 1.1.3; внутренние проверки не перечисляются.
$rules=[
 ['Готов к обновлению','К обновлению','Артикул Excel однозначно найден на сайте. GUID Excel не занят другим товаром; цена допустима, остальные проверки пройдены. XML_ID сайта может быть равен 0 или содержать другой GUID.','Записать GUID Excel в XML_ID сайта, обновить цену и остальные предусмотренные поля.'],
 ['GUID уже совпадает','К обновлению','Товар однозначно найден по артикулу, XML_ID сайта совпадает с GUID Excel; остальные проверки пройдены.','Оставить GUID без изменений, обновить цену и остальные предусмотренные поля.'],
 ['Дубль артикула в Excel','Пропустить','Один артикул встречается в нескольких строках проверяемого Excel. GUID проверяется для диагностики.','Пропустить все строки с этим артикулом, показать найденные товары для проверки.'],
 ['Дубль артикула на сайте','Пропустить','По артикулу Excel найдено несколько товаров или торговых предложений в Битрикс.','Не выбирать товар автоматически, показать найденные ID для ручной проверки.'],
 ['GUID найден у другого артикула на сайте','Пропустить','По артикулу Excel товар не найден, но по GUID обнаружен товар с другим артикулом. Дополнительно проверяются скрытые различия артикулов и возможное преобразование артикула Excel в дату.','Показать найденный товар и его ID, обнаруженные различия и предупреждение о дате (если есть). Ничего не обновлять.'],
 ['GUID занят другим товаром на сайте','Пропустить','По артикулу найден один товар, но GUID Excel уже принадлежит другому товару сайта.','Показать конфликтующие товары, ничего не менять.'],
 ['Товар не найден на сайте','Пропустить','Артикул в Excel заполнен, но товар не найден ни по артикулу, ни по однозначному совпадению GUID. Дополнительно проверяется возможное преобразование артикула Excel в дату.','Пропустить строку; новые товары не создавать. При обнаружении даты показать причину.'],
 ['Нет артикула в Excel','Пропустить','В Excel пустой артикул. GUID всё равно проверяется для диагностики.','Пропустить независимо от результата поиска по GUID.'],
 ['Проблема с ценой в Excel','Пропустить','Цена отсутствует, равна нулю либо не распознана как допустимое число. Если одновременно есть дубль артикула, основным статусом будет дубль.','Не обновлять товар; в причинах показать: «Пустая цена», «Цена 0» или «Некорректная цена».']
];
// Локальные стили инструмента: не затрагивают остальные страницы админки.
echo '<style>
.bgt-guide{max-width:1250px;margin:14px 0}.bgt-guide summary{cursor:pointer;padding:9px 11px;font-size:14px;background:#f3f6f8;border:1px solid #dce3e9;border-radius:5px}.bgt-guide[open] summary{border-radius:5px 5px 0 0}.bgt-guide-content{padding:14px;border:1px solid #dce3e9;border-top:0;border-radius:0 0 5px 5px;line-height:1.55}.bgt-guide-content li{margin:0 0 13px}.bgt-guide-content li:last-child{margin-bottom:0}.bgt-guide-content p{margin:5px 0}.bgt-rules-wrap{overflow-x:auto}.bgt-rules{border-collapse:collapse;width:100%;table-layout:fixed;font-size:13px}.bgt-rules th,.bgt-rules td{padding:12px 11px;border:1px solid #dce3e9;vertical-align:top;text-align:left;overflow-wrap:anywhere;word-break:normal;white-space:normal;line-height:1.5}.bgt-rules th{background:#eaf0f5}.bgt-rules tbody tr:nth-child(even){background:#f6f8fa}.bgt-rules .bgt-ready{color:#23833b;font-weight:600}.bgt-rules .bgt-skip{color:#666;font-weight:600}
.bgt-actions{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin:10px 0}.bgt-actions form{display:inline-flex;margin:0}.bgt-stage{max-width:1250px;margin:17px 0}.bgt-stage h3{margin-bottom:7px}.bgt-stage-note{color:#667480}.bgt-separator{border:0;border-top:1px solid #e1e6ea;margin:18px 0}
@media(max-width:760px){.bgt-rules{min-width:760px}.bgt-guide-content{padding:10px}}
</style>';
echo '<details class="bgt-guide"><summary><strong>Порядок подготовки и сверки Excel с Битрикс</strong></summary><div class="bgt-guide-content">';
echo '<p>Рекомендуется выполнять этапы по порядку. До запуска обновления изменения на сайте не производятся.</p><ol>';
echo '<li><strong>Исключить товары без цены.</strong><p>Удалить из рабочего Excel строки с пустой или нулевой ценой. Это можно сделать <strong>вручную</strong> перед загрузкой или <strong>через подготовку</strong>: выбрать «Исключить строки с нулевой или пустой ценой (если не удалены вручную из Excel)» и нажать «Применить подготовку».</p><p><strong>Результат:</strong> подготовленный Excel с товарами, у которых есть цена.</p></li>';
echo '<li><strong>Проверить и устранить дубли артикулов в Excel.</strong><p>После исключения строк без цены скачайте «Дубли XLSX», проверьте повторяющиеся артикулы и по возможности исправьте исходные данные до сверки. При необходимости повторно загрузите исправленный Excel. Если дубли останутся, сверка всё равно выполнится, но все строки с повторяющимися артикулами получат статус «Дубль артикула в Excel» и будут пропущены при обновлении.</p><p><strong>Результат:</strong> желательно получить Excel без дублей, чтобы больше товаров могли пройти сверку и обновление.</p></li>';
echo '<li><strong>Проверить артикулы Excel на преобразование в даты.</strong><p>Нажать «Проверить артикулы Excel». При обнаружении дат проверить и исправить исходный XLSX. Артикулы автоматически не восстанавливаются; строки с датами будут пропущены при обновлении даже без нажатия этой кнопки.</p><p><strong>Результат:</strong> возможные повреждения артикулов выявлены до сверки.</p></li>';
echo '<li><strong>Выполнить предварительную сверку с Битрикс.</strong><p>Выбрать профиль сайта и запустить сверку. Программа ищет товары по артикулу, проверяет принадлежность GUID и выявляет конфликты. На этом этапе данные сайта не изменяются.</p><p><strong>Результат:</strong> сводка по статусам и журнал проблемных строк; полный список доступен в XLSX.</p></li>';
echo '<li><strong>Проверить результаты сверки.</strong><p>Посмотреть сводку по 9 статусам и журнал проблемных строк с фильтром. При необходимости скачать предварительный XLSX, исправить проблемы и повторить сверку. Особое внимание уделить дублям на сайте и конфликтам GUID.</p><p><strong>Результат:</strong> проверенный список товаров, которые разрешено обновлять.</p></li>';
echo '<li><strong>Запустить обновление.</strong><p>После проверки результатов обновлять разрешённые товары партиями по 25 строк. По завершении скачать итоговый XLSX.</p><p><strong>Результат:</strong> обновлённые товары, итоговый отчёт и сохранённая история операции.</p></li>';
echo '</ol></div></details>';
echo '<details class="bgt-guide"><summary><strong>Логика сверки по каждому статусу</strong></summary><div class="bgt-guide-content">';
echo '<p>Если у строки несколько проблем, статус отражает приоритетную причину; остальные проверки отображаются в колонке «Причины». Числовой артикул — предупреждение, а не самостоятельный запрет.</p>';
echo '<div class="bgt-rules-wrap"><table class="bgt-rules"><colgroup><col style="width:21%"><col style="width:13%"><col style="width:34%"><col style="width:32%"></colgroup><thead><tr><th>Статус</th><th>Результат</th><th>Условия сверки</th><th>Действие</th></tr></thead><tbody>';
foreach($rules as $i=>$rule){echo '<tr><td><strong>'.($i+1).'. '.h($rule[0]).'</strong></td><td class="'.($rule[1]==='К обновлению'?'bgt-ready':'bgt-skip').'">'.h($rule[1]).'</td><td>'.h($rule[2]).'</td><td>'.h($rule[3]).'</td></tr>';}
echo '</tbody></table></div><p><strong>Дополнительная проверка артикулов Excel:</strong> если артикул сохранён как дата, строка всегда пропускается независимо от статуса. В колонке «Причины» отображается предупреждение. Артикулы автоматически не исправляются.</p></div></details>';
echo '<details class="bgt-guide"><summary><strong>История изменений по версиям</strong></summary><div class="bgt-guide-content">';
$changelog=dirname(__DIR__,2).'/local/bitrix-guid-tool/CHANGELOG.md';
if(is_readable($changelog)){
 $lines=file($changelog,FILE_IGNORE_NEW_LINES);$open=false;
 foreach($lines as $line){
  if(preg_match('/^## \[([0-9]+\.[0-9]+\.[0-9]+)\](.*)$/u',$line,$m)){
   if($open)echo '</ul></details>';
   echo '<details'.($m[1]==='1.1.5'?' open':'').'><summary><strong>'.h($m[1]).'</strong>'.h($m[2]).'</summary><ul>';$open=true;
  }elseif($open&&preg_match('/^## /',$line)){echo '</ul></details>';$open=false;}
  elseif($open&&preg_match('/^- (.+)$/u',$line,$m)){echo '<li>'.h($m[1]).'</li>';}
 }
 if($open)echo '</ul></details>';
}else echo '<p>CHANGELOG.md недоступен.</p>';
echo '</div></details>';
echo '<form method="get"><input type="hidden" name="tab_token" value="'.h($tabToken).'">Сайт: <select name="profile"><option value="">Выберите сайт</option>';foreach(['gk-mact.ru','prom-mact.ru'] as $s)echo '<option value="'.h($s).'"'.($profile===$s?' selected':'').'>'.h($s).'</option>';echo '</select> <button>Открыть</button></form>';
if($cfg){echo '<p>Каталог '.$cfg['products'].' / SKU '.$cfg['offers'].'. Этап: '.h($job['phase']??'нет задания').'</p>';
if(!$job||in_array($job['phase']??'',['finished','stopped'],true)){form('upload',$profile);echo '<p><input type="file" name="xlsx" accept=".xlsx" required> <button>Загрузить Excel</button></p></form>';}
if(in_array($job['phase']??'',['uploaded','preview','prepared'],true)){
    $d=$job['data'];$width=$d['width'];echo '<h3>Excel: подготовка (необязательно) или сразу проверка</h3><div style="overflow:auto;max-height:250px"><table class="adm-list-table"><tr>';for($i=0;$i<$width;$i++)echo '<th>'.h(X::col($i)).'</th>';echo '</tr>';foreach(array_slice($d['rows'],0,8,true) as $n=>$r){echo '<tr>';for($i=0;$i<$width;$i++)echo '<td>'.h($r[$i]['v']??'').'</td>';echo '</tr>';}echo '</table></div>';
    form('prepare_preview',$profile);foreach(['article'=>'Артикул','guid'=>'GUID','name'=>'Название','price'=>'Цена'] as $k=>$label){echo '<p>'.h($label).' <select name="'.$k.'">';for($i=0;$i<$width;$i++){$header=$d['rows'][1][$i]['v']??X::col($i);$selected=isset($job['map'][$k])&&$job['map'][$k]===$i?' selected':'';echo '<option value="'.$i.'"'.$selected.'>'.h(X::col($i).' — '.$header).'</option>';}echo '</select></p>';}
    echo '<p>Первая строка данных <input type="number" name="start" min="2" value="'.(int)($job['map']['start']??2).'" required></p>';
    foreach(['drop_price'=>'Исключить строки с нулевой или пустой ценой (если не удалены вручную из Excel)','normalize'=>'Преобразовать артикулы в текст в новом XLSX (необязательно)'] as $k=>$label){$on=isset($job['opts'])?!empty($job['opts'][$k]):($k==='drop_price');echo '<p><label><input type="checkbox" name="'.$k.'" value="1"'.($on?' checked':'').'> '.h($label).'</label></p>'; }
    echo '<p>Пробелы по краям артикула убираются при поиске всегда. Числовой артикул — предупреждение, не причина пропуска.</p>';
    echo '<div class="bgt-stage"><h3>1. Подготовка Excel</h3><div class="bgt-actions"><button name="action" value="check_dates">Проверить артикулы Excel</button><button  name="action" value="prepare_preview">Проверить подготовку</button><button class="adm-btn-save" name="action" value="prepare_apply">Применить подготовку</button></div><p class="bgt-stage-note">Альтернативный вариант — сверка исходного Excel без подготовки</p><div class="bgt-actions"><button  name="action" value="check_direct">Проверить исходный Excel без подготовки</button></div></div></form>';
    if(array_key_exists('date_check',$job)){
        $dateRows=$job['date_check'];echo '<div class="bgt-stage"><h3>Предварительная проверка артикулов Excel на даты</h3><p>Найдено проблемных строк: '.count($dateRows).'.</p>';
        if($dateRows){echo '<div style="overflow:auto;max-height:300px"><table class="adm-list-table"><tr><th>Строка Excel</th><th>Причина</th></tr>';foreach($dateRows as $item)echo '<tr><td>'.(int)$item['n'].'</td><td>'.h($item['issue']).'</td></tr>';echo '</table></div><p>Исправьте исходный XLSX и загрузите заново. Иначе строки будут пропущены при обновлении.</p>';}
        else echo '<p>Ячейки с форматом даты не обнаружены.</p>';echo '</div>';
    }
    if(isset($job['stats'])){
        $st=$job['stats'];echo '<h3>Статистика исходного Excel</h3><p>Всего строк: '.(int)$st['total'].'; цена 0 или пустая: '.(int)$st['zero_price'].'; пустой артикул: '.(int)$st['empty_article'].'; пробелы по краям: '.(int)$st['normalized'].'.</p>';
        echo '<p>По выбранным настройкам будет исключено: '.((int)$st['total']-(int)$st['kept']).'; останется: '.(int)$st['kept'].'.</p>';
        $dups=C::duplicates($job['rows']);echo '<div class="bgt-stage"><h3>2. Найденные дубли артикулов</h3><p>Дубли определены по оставшимся после подготовки строкам. Дубли не удаляются и не изменяются. Файл «Дубли XLSX» содержит все исходные столбцы в исходном порядке.</p>'; 
        echo '<p>Строк с повторяющимися артикулами: '.count($dups).'.</p>';
        if($dups){form('duplicates_xlsx',$profile);echo '<div class="bgt-actions"><button >Скачать дубли XLSX</button></div></form>';}
        else echo '<p>Дубли не найдены.</p>'; echo '</div>';

    }
}
if(($job['phase']??'')==='prepared'){echo '<div class="bgt-stage"><h3>3. Подготовленный файл</h3><p class="bgt-stage-note">Подготовка применена. Сверка будет выполнена только по строкам, оставшимся после подготовки.</p>';form('prepared_xlsx',$profile);echo '<div class="bgt-actions"><button >Скачать подготовленный XLSX</button></div></form>';form('plan',$profile);echo '<div class="bgt-actions"><button class="adm-btn-save">Полная сверка с Битрикс без записи</button></div></form></div>';}
if(in_array($job['phase']??'',['planned','processing','stopped','finished'],true)){
    $s=S::count($job);echo '<h3>Сверка и обновление</h3><p><strong>К обновлению: '.$s['ready'].'; Будут пропущены: '.$s['will_skip'].'</strong> (по строкам сверки)</p>';
    echo '<div id="bgt-progress" aria-live="polite"><p>Всего <span id="bgt-total">'.$s['total'].'</span>; успешно <span id="bgt-success">'.$s['success'].'</span>; ошибок <span id="bgt-errors">'.$s['errors'].'</span>; пропущено <span id="bgt-skipped">'.$s['skipped'].'</span>; осталось <span id="bgt-remaining">'.$s['remaining'].'</span></p>';
    $done=$s['success']+$s['errors'];$pct=$s['ready']?min(100,round($done*100/$s['ready'],1)):100;
    echo '<progress id="bgt-progressbar" max="100" value="'.$pct.'" style="width:350px;max-width:100%"></progress> <span id="bgt-percent">'.$pct.'%</span><p id="bgt-batch-message" role="status"></p></div>';
    foreach(['prepared_xlsx'=>'Подготовленный XLSX','preliminary_xlsx'=>'Предварительный XLSX','final_xlsx'=>'Итоговый XLSX'] as $act=>$title){if($act==='prepared_xlsx'&&empty($job['prepared_applied']))continue;if($act==='final_xlsx'&&$job['phase']==='planned')continue;form($act,$profile);if($act==='final_xlsx')echo '<span id="bgt-final-download">';echo '<button>'.h($title).'</button>';if($act==='final_xlsx')echo '</span>';echo '</form>';}
    if($job['phase']==='planned'){form('final_xlsx',$profile);echo '<span id="bgt-final-download-new" hidden><button>Итоговый XLSX</button></span></form>';}
    if(in_array($job['phase'],['planned','processing'],true)&&$s['remaining']>0){form('batch',$profile);echo '<button type="submit" id="bgt-batch-btn" class="adm-btn-save">Обновить следующие '.min(25,$s['remaining']).'</button></form>';}
    echo '<p id="bgt-completed"'.($job['phase']==='finished'?'':' hidden').'><strong>Обновление завершено.</strong> Итоговый XLSX сохранён в истории.</p>';
    // Сводка по девяти статусам; журнал по умолчанию показывает проблемы, выбор статуса — все строки этого статуса.
    $all=$job['rows'];$statusCounts=[];
    foreach($all as $r){$st=(string)($r['status']??'');$statusCounts[$st]=($statusCounts[$st]??0)+1;}
    echo '<div class="bgt-stage"><h3>Сводка по статусам</h3><table class="adm-list-table"><tr><th>Статус</th><th>Количество строк</th></tr>';
    foreach($rules as $rule){$st=$rule[0];echo '<tr><td>'.h($st).'</td><td>'.(int)($statusCounts[$st]??0).'</td></tr>';}
    echo '</table></div>';
    $q=trim((string)($_GET['q']??''));$filter=(string)($_GET['status']??'');
    $problemRows=array_values(array_filter($all,function($r){
        // The original decision is retained after batch updates; completed rows must not become conflicts.
        $decision=(string)($r['initial_result']??$r['result']??'');
        if($decision==='Пропустить')return true;
        if($decision==='К обновлению')return false;
        // Compatibility with older saved jobs without initial_result.
        return !in_array($r['status']??'',['Готов к обновлению','GUID уже совпадает'],true)||!empty($r['date_issue']);
    }));
    // Для «Все проблемы» учитываем исходный допуск; для выбранного статуса — все строки с этим статусом.
    $filtered=array_values(array_filter($filter===''?$problemRows:$all,function($r)use($q,$filter){
        if($filter!==''&&($r['status']??'')!==$filter)return false;
        if($q==='')return true;
        $hay=implode(' ',[$r['article']??'',$r['excel_name']??'',$r['guid']??'',$r['id']??'',$r['name']??'',$r['status']??'',$r['issues']??'']);
        return function_exists('mb_stripos')?mb_stripos($hay,$q,0,'UTF-8')!==false:stripos($hay,$q)!==false;
    }));
    $limit=200;
    echo '<div class="bgt-stage"><h3>Журнал сверки</h3>';
    echo '<form method="get"><input type="hidden" name="tab_token" value="'.h($tabToken).'"><input type="hidden" name="profile" value="'.h($profile).'">Поиск <input name="q" value="'.h($q).'"> Статус <select name="status"><option value="">Все проблемы</option>';
    foreach($rules as $rule){$st=$rule[0];echo '<option value="'.h($st).'"'.($st===$filter?' selected':'').'>'.h($st).'</option>';}
    echo '</select> <button>Найти</button></form>';
    echo '<p>Всего проблемных строк: '.count($problemRows).'; по выбранному фильтру: '.count($filtered).'. Показано не более '.$limit.'. Полный список — в предварительном XLSX.</p>';
    if($filtered){echo '<div style="overflow:auto;max-height:650px"><table class="adm-list-table"><tr><th>Строка Excel</th><th>Артикул Excel</th><th>GUID Excel</th><th>ID сайта</th><th>Артикул/название сайта</th><th>Статус</th><th>Причины</th></tr>';
        foreach(array_slice($filtered,0,$limit) as $r){echo '<tr>';foreach([$r['n']??'',$r['article']??'',$r['guid']??'',$r['id']??'',$r['name']??'',$r['status']??'',$r['issues']??''] as $v)echo '<td>'.h($v).'</td>';echo '</tr>';}
        echo '</table></div>';
    }else echo '<p>Строк по выбранному фильтру нет.</p>';
    echo '</div>';

}
if($job){echo '<div class="bgt-stage"><hr class="bgt-separator"><h3>4. Управление проверкой</h3>';form('reset',$profile);echo '<div class="bgt-actions"><button  onclick="return confirm(\'Сбросить текущее задание? Архив останется.\')">Новая проверка / сброс</button></div></form></div>';}
echo <<<'BGTJS'
<script>
(function(){
  'use strict';
  const key='bgt-tab-token';
  let token=sessionStorage.getItem(key);
  if(!token){const bytes=new Uint8Array(16);crypto.getRandomValues(bytes);token=Array.from(bytes,b=>b.toString(16).padStart(2,'0')).join('');sessionStorage.setItem(key,token);}
  const url=new URL(location.href);
  if(url.searchParams.get('tab_token')!==token){url.searchParams.set('tab_token',token);location.replace(url.toString());return;}
  document.querySelectorAll('input[name="tab_token"]').forEach(el=>el.value=token);
  const btn=document.getElementById('bgt-batch-btn');if(!btn)return;
  const form=btn.closest('form');let busy=false;let expectedDone=Number(document.getElementById('bgt-success').textContent)+Number(document.getElementById('bgt-errors').textContent);
  form.addEventListener('submit',async function(event){
    event.preventDefault();if(busy)return;
    busy=true;btn.disabled=true;const message=document.getElementById('bgt-batch-message');message.textContent='Обрабатывается партия до 25 товаров…';
    try{
      const body=new FormData(form);body.set('ajax','1');body.set('tab_token',token);body.set('expected_done',String(expectedDone));
      const response=await fetch(location.pathname,{method:'POST',body,credentials:'same-origin',cache:'no-store',headers:{'Accept':'application/json'}});
      const data=await response.json();if(!response.ok||!data.ok)throw new Error(data.error||'Ошибка обновления');
      const s=data.stats;expectedDone=s.success+s.errors;for(const field of ['total','success','errors','skipped','remaining']){document.getElementById('bgt-'+field).textContent=s[field];}
      const pct=s.ready?Math.min(100,Math.round((s.success+s.errors)*1000/s.ready)/10):100;
      document.getElementById('bgt-progressbar').value=pct;document.getElementById('bgt-percent').textContent=pct+'%';
      message.textContent=data.message||'Партия обработана';
      const dl=document.getElementById('bgt-final-download-new');if(dl&&data.archive_ready)dl.hidden=false;
      if(s.remaining===0){btn.hidden=true;document.getElementById('bgt-completed').hidden=false;}
      else{btn.textContent='Обновить следующие '+Math.min(25,s.remaining);}
    }catch(err){message.textContent='Ошибка: '+err.message+'. Проверьте статистику перед повтором. При неизвестном результате запроса не нажимайте повторно без проверки.';}
    finally{busy=false;btn.disabled=false;}
  });
})();
</script>
BGTJS;
echo '<h3>История обновлений</h3><p>Отчёты хранятся бессрочно до ручного удаления.</p>';
$archive=$store->archiveList($profile);if(!$archive)echo '<p>Архив пока пуст.</p>';else{
    form('archive_delete',$profile);echo '<input type="hidden" name="confirm" value="yes"><table class="adm-list-table"><tr><th>Выбор</th><th>Начало</th><th>Завершение</th><th>Сайт</th><th>Администратор</th><th>Этап</th><th>Обновлено</th><th>Ошибок</th><th>Отчёт</th></tr>';
    foreach($archive as $a){echo '<tr><td><input type="checkbox" name="ids[]" value="'.h($a['id']).'"></td><td>'.h($a['started']).'</td><td>'.h($a['completed']).'</td><td>'.h($a['profile']).'</td><td>'.(int)$a['admin'].'</td><td>'.h($a['phase']).'</td><td>'.(int)($a['stats']['success']??0).'</td><td>'.(int)($a['stats']['errors']??0).'</td><td><button type="submit" form="download-'.h($a['id']).'">Скачать XLSX</button></td></tr>';}
    echo '</table><button onclick="return confirm(\'Безвозвратно удалить выбранные архивные отчёты? Товары Битрикс не изменятся.\')">Удалить выбранные</button></form>';
    foreach($archive as $a){echo '<form id="download-'.h($a['id']).'" method="post" style="display:none">'.bitrix_sessid_post().'<input type="hidden" name="profile" value="'.h($profile).'"><input type="hidden" name="action" value="archive_download"><input type="hidden" name="id" value="'.h($a['id']).'"></form>';}
}
}
require $_SERVER['DOCUMENT_ROOT'].'/bitrix/modules/main/include/epilog_admin.php';

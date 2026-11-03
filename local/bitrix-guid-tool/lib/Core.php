<?php
namespace GuidTool;
final class Core {
    public static function trim($s){return preg_replace('/^[\s\x{00A0}]+|[\s\x{00A0}]+$/u','',(string)$s);}
    public static function key($s){$s=self::trim($s);return function_exists('mb_strtolower')?mb_strtolower($s,'UTF-8'):strtolower($s);}
    public static function price($s){$s=self::trim($s);if($s==='')return [null,'Пустая цена'];$s=str_replace(["\xc2\xa0","\xe2\x80\xaf"],' ',$s);if(!preg_match('/^(?:\d+|\d{1,3}(?: \d{3})+)(?:[.,]\d{1,2})?$/D',$s))return [null,'Некорректная цена'];$s=str_replace([' ', ','],['','.'],$s);if(!is_numeric($s)||!is_finite((float)$s))return [null,'Некорректная цена'];if((float)$s<=0)return [null,'Цена 0'];return [number_format((float)$s,2,'.',''),''];}
    public static function config($profile){$all=['gk-mact.ru'=>[384,386],'prom-mact.ru'=>[488,490]];if(!isset($all[$profile]))throw new \RuntimeException('Выберите сайт');return ['domain'=>$profile,'products'=>$all[$profile][0],'offers'=>$all[$profile][1]];}
    public static function validate($cfg){$link=\CCatalogSKU::GetInfoByProductIBlock($cfg['products']);if(!$link||(int)$link['IBLOCK_ID']!==$cfg['offers']||empty($link['SKU_PROPERTY_ID']))throw new \RuntimeException('Не подтверждена связь SKU');foreach([$cfg['products']=>'CML2_ARTICLE',$cfg['offers']=>'ARTICLE'] as $ib=>$code){$p=\CIBlockProperty::GetList([],['IBLOCK_ID'=>$ib,'CODE'=>$code])->Fetch();if(!$p||$p['PROPERTY_TYPE']!=='S'||$p['MULTIPLE']!=='N')throw new \RuntimeException('Неверное свойство артикула '.$code);}self::kod($cfg['products']);self::kod($cfg['offers']);return (int)$link['SKU_PROPERTY_ID'];}
    public static function kod($ib){static $cache=[];if(isset($cache[$ib]))return $cache[$ib];$a=[];$rs=\CIBlockProperty::GetList([],['IBLOCK_ID'=>$ib]);while($p=$rs->Fetch())if($p['CODE']==='KOD_IZ_1C'||$p['XML_ID']==='KOD_IZ_1C')$a[$p['ID']]=$p;if(count($a)!==1)throw new \RuntimeException('Свойство KOD_IZ_1C не определено: '.$ib);$p=reset($a);if($p['PROPERTY_TYPE']!=='L'||$p['MULTIPLE']!=='N')throw new \RuntimeException('KOD_IZ_1C должно быть одиночным списком');$yes=[];$rs=\CIBlockPropertyEnum::GetList([],['PROPERTY_ID'=>$p['ID']]);while($v=$rs->Fetch())if($v['XML_ID']==='Y')$yes[]=(int)$v['ID'];if(count($yes)!==1)throw new \RuntimeException('Не найден единственный вариант Y для KOD_IZ_1C');return $cache[$ib]=['id'=>(int)$p['ID'],'yes'=>$yes[0]];}
    public static function catalog($cfg,$link){$items=[];$parents=[];$owners=[];$index=[];foreach([$cfg['offers']=>'ARTICLE',$cfg['products']=>'CML2_ARTICLE'] as $ib=>$code){$k=self::kod($ib);$sel=['ID','IBLOCK_ID','NAME','XML_ID','ACTIVE','PROPERTY_'.$code,'PROPERTY_'.$k['id']];if($ib===$cfg['offers'])$sel[]='PROPERTY_'.$link;$rs=\CIBlockElement::GetList(['ID'=>'ASC'],['IBLOCK_ID'=>$ib,'CHECK_PERMISSIONS'=>'N'],false,false,$sel);while($e=$rs->Fetch()){$id=(int)$e['ID'];if($ib===$cfg['offers'] && (int)$e['PROPERTY_'.$link.'_VALUE']>0)$parents[(int)$e['PROPERTY_'.$link.'_VALUE']]=true;$items[$id]=['id'=>$id,'ib'=>$ib,'article'=>self::trim($e['PROPERTY_'.$code.'_VALUE']),'name'=>(string)$e['NAME'],'guid'=>(string)$e['XML_ID'],'active'=>(string)$e['ACTIVE'],'kod'=>(int)($e['PROPERTY_'.$k['id'].'_ENUM_ID']??$e['PROPERTY_'.$k['id'].'_VALUE']??0)];if((string)$e['XML_ID']!=='')$owners[self::key($e['XML_ID'])][$id]=true;}}
        $ids=array_keys($items);foreach(array_chunk($ids,500) as $chunk){$rs=\Bitrix\Catalog\ProductTable::getList(['select'=>['ID','QUANTITY'],'filter'=>['@ID'=>$chunk]]);while($p=$rs->fetch())$items[(int)$p['ID']]['quantity']=(string)$p['QUANTITY'];}
        $base=\Bitrix\Catalog\GroupTable::getList(['select'=>['ID'],'filter'=>['=BASE'=>'Y']])->fetch();if(!$base)throw new \RuntimeException('Не найдена базовая группа цен');$prices=[];foreach(array_chunk($ids,500) as $chunk){$rs=\Bitrix\Catalog\PriceTable::getList(['select'=>['ID','PRODUCT_ID','PRICE','CURRENCY','QUANTITY_FROM','QUANTITY_TO'],'filter'=>['@PRODUCT_ID'=>$chunk,'=CATALOG_GROUP_ID'=>(int)$base['ID']]]);while($p=$rs->fetch()){if(($p['QUANTITY_FROM']!==null&&(float)$p['QUANTITY_FROM']>1)||($p['QUANTITY_TO']!==null&&(float)$p['QUANTITY_TO']<1))continue;$id=(int)$p['PRODUCT_ID'];if(array_key_exists($id,$prices))$prices[$id]=['ambiguous'=>true];else $prices[$id]=['id'=>(int)$p['ID'],'price'=>(string)$p['PRICE'],'currency'=>(string)$p['CURRENCY']];}}
        foreach($items as $id=>&$e){$e['price']=$prices[$id]??null;if($e['ib']===$cfg['products']&&isset($parents[$id]))continue;if($e['article']!=='')$index[self::key($e['article'])][$id]=$e;}unset($e);return [$index,$owners,$items,(int)$base['ID']];
    }
    public static function dateIssue($cell){
        if(empty($cell['date']))return '';
        $value=(string)($cell['date_value']??'');$raw=(string)($cell['v']??'');
        $suffix=($raw!==''&&is_numeric($raw))?' (числовое значение '.$raw.')':'';
        return 'Артикул Excel сохранён как дата: '.($value!==''?$value:$raw).$suffix.'. Возможно автоматическое преобразование артикула Excel в дату.';
    }
    public static function articleDifferences($excel,$site){
        // Diagnostic-only canonicalization; never used for product matching.
        $conf=['А'=>'A','В'=>'B','Е'=>'E','К'=>'K','М'=>'M','Н'=>'H','О'=>'O','Р'=>'P','С'=>'C','Т'=>'T','Х'=>'X','а'=>'a','е'=>'e','о'=>'o','р'=>'p','с'=>'c','у'=>'y','х'=>'x'];
        $canonical=function($value)use($conf){
            $value=preg_replace('/[\x{200B}-\x{200F}\x{2060}\x{FEFF}]/u','',(string)$value);
            $value=str_replace(["\xc2\xa0","\xe2\x80\xaf"],' ',$value);
            $value=preg_replace('/ +/u',' ',$value);
            return strtr(trim($value),$conf);
        };
        if($canonical($excel)!==$canonical($site))return [];
        $messages=[];
        preg_match_all('/[ ]+/u',(string)$excel,$ex,PREG_OFFSET_CAPTURE);preg_match_all('/[ ]+/u',(string)$site,$si,PREG_OFFSET_CAPTURE);
        $a=array_map(function($v){return strlen($v[0]);},$ex[0]);$b=array_map(function($v){return strlen($v[0]);},$si[0]);
        if($a!==$b)$messages[]='Количество обычных пробелов отличается: Excel — '.implode(', ',$a?:[0]).'; Битрикс — '.implode(', ',$b?:[0]);
        if(preg_match('/^[ ]|[ ]$/u',$excel)||preg_match('/^[ ]|[ ]$/u',$site))$messages[]='Проверьте пробелы в начале или конце артикула';
        $chars=function($s){return preg_split('//u',(string)$s,-1,PREG_SPLIT_NO_EMPTY);};$x=$chars($excel);$y=$chars($site);
        $normalize=function($v)use($conf){return strtr($v,$conf);};
        if($x!==$y&&$normalize($excel)===$normalize($site))$messages[]='Есть похожие русские и латинские буквы';
        if(preg_match('/[\x{00A0}\x{202F}]/u',$excel.$site))$messages[]='Обнаружен неразрывный пробел (отличается от обычного)';
        if(preg_match('/[\x{200B}-\x{200F}\x{2060}\x{FEFF}]/u',$excel.$site))$messages[]='Обнаружены невидимые символы';
        return $messages;
    }
    public static function prepare($data,$map,$opts){
        $out=[];$stats=['total'=>0,'excluded_price'=>0,'excluded_article'=>0,'empty_article'=>0,'zero_price'=>0,'normalized'=>0,'duplicates'=>0,'kept'=>0];
        foreach($data['rows'] as $n=>$cells){
            if($n<$map['start'])continue;
            $stats['total']++;
            $article=(string)($cells[$map['article']]['v']??'');$norm=self::trim($article);
            if($norm!==$article)$stats['normalized']++;
            if($norm==='')$stats['empty_article']++;
            [$price,$priceError]=self::price($cells[$map['price']]['v']??'');
            if(in_array($priceError,['Пустая цена','Цена 0'],true))$stats['zero_price']++;
            $exclude=!empty($opts['drop_price'])&&in_array($priceError,['Пустая цена','Цена 0'],true);
            if($exclude)$stats['excluded_price']++;else $stats['kept']++;
            $out[]=['n'=>$n,'cells'=>$cells,'article'=>$norm,'excel_name'=>$cells[$map['name']]['v']??'',
                'date_issue'=>self::dateIssue($cells[$map['article']]??[]),'guid'=>(string)($cells[$map['guid']]['v']??''),'price'=>$price,'price_error'=>$priceError,
                'exclude'=>$exclude?'Исключён: '.$priceError:'','status'=>'','issues'=>'','result'=>'К обновлению',
                'id'=>0,'ib'=>0,'name'=>'','old_guid'=>'','site_price'=>'','type'=>''];
        }
        $stats['duplicates']=count(self::duplicates($out));
        return [$out,$stats];
    }
    public static function duplicates($rows){
        $counts=[];foreach($rows as $r){if(!empty($r['exclude']))continue;$key=self::key($r['article']);if($key!=='')$counts[$key]=($counts[$key]??0)+1;}
        $out=[];foreach($rows as $r){if(!empty($r['exclude']))continue;$key=self::key($r['article']);if($key!==''&&($counts[$key]??0)>1)$out[]=$r;}
        usort($out,function($a,$b){$cmp=strcmp(self::key($a['article']),self::key($b['article']));return $cmp?:($a['n']<=>$b['n']);});return $out;
    }
    public static function plan(&$job,$cfg,$link){
        [$index,$owners,$items,$base]=self::catalog($cfg,$link);
        $counts=[];$guidCounts=[];
        foreach($job['rows'] as $r){if(!empty($r['exclude']))continue;$a=self::key($r['article']);$g=self::key($r['guid']);if($a!=='')$counts[$a]=($counts[$a]??0)+1;if($g!=='')$guidCounts[$g][$a]=true;}
        foreach($job['rows'] as &$r){
            if(!empty($r['exclude'])){$r['status']='Исключён при подготовке';$r['issues']=$r['exclude'];$r['result']='Пропустить';continue;}
            $a=self::key($r['article']);$g=self::key($r['guid']);$m=$job['map'];$cell=$r['cells'];$issues=[];$warnings=[];$status='';
            if(!empty($cell[$m['article']]['error']))$issues[]='Артикул: '.$cell[$m['article']]['error'];
            if(!empty($cell[$m['guid']]['error']))$issues[]='GUID: '.$cell[$m['guid']]['error'];
            if(!empty($cell[$m['price']]['error']))$issues[]='Цена: '.$cell[$m['price']]['error'];
            if(!empty($r['date_issue']))$issues[]=$r['date_issue'];
            if(!empty($cell[$m['article']]['numeric'])&&empty($r['date_issue']))$warnings[]='Числовой артикул: проверьте возможную потерю ведущих нулей';
            if(!empty($cell[$m['guid']]['numeric']))$issues[]='GUID должен быть текстом';
            if($g==='')$issues[]='Не пройдена проверка обязательных данных';
            if($r['price_error'])$issues[]=$r['price_error'];
            if($g!==''&&count($guidCounts[$g]??[])>1)$issues[]='Один GUID у разных артикулов Excel';
            $matches=$a!==''?($index[$a]??[]):[];
            if(count($matches)===1){
                $e=reset($matches);$r['id']=$e['id'];$r['ib']=$e['ib'];$r['name']=$e['name'];$r['old_guid']=$e['guid'];
                $r['site_price']=$e['price']['price']??'';$r['type']=$e['ib']===$cfg['offers']?'SKU':'Простой товар';$r['snapshot']=$e;
                if(!array_key_exists('quantity',$e))$issues[]='Нет записи каталога';
                if(isset($e['price']['ambiguous']))$issues[]='Несколько базовых цен для количества 1';
            }
            $occupied=$g!==''?($owners[$g]??[]):[];
            if($r['id'])unset($occupied[$r['id']]);
            $guidMatch=null;
            if($occupied){
                $ids=array_keys($occupied);
                if(count($matches)>0)$issues[]='GUID занят другим товаром: '.implode(', ',$ids);
                if(count($matches)===0&&count($ids)===1&&isset($items[$ids[0]])){
                    $guidMatch=$items[$ids[0]];
                    if(!empty($r['date_issue']))$issues[]='GUID найден у товара ID '.$guidMatch['id'].' с артикулом '.$guidMatch['article'].'.';
                    else{$issues[]='GUID найден у товара ID '.$guidMatch['id'].'; Артикул Excel: '.$r['article'].'; артикул Битрикс: '.$guidMatch['article'];foreach(self::articleDifferences($r['article'],$guidMatch['article']) as $difference)$issues[]=$difference;}
                    $r['id']=$guidMatch['id'];$r['name']=$guidMatch['name'];$r['old_guid']=$guidMatch['guid'];$r['type']=$guidMatch['ib']===$cfg['offers']?'SKU':'Простой товар';
                }
            }
            // Priority: duplicates, conflicting GUID ownership, matching diagnostics, then price.
            if($a==='')$status='Нет артикула в Excel';
            elseif(($counts[$a]??0)>1){$status='Дубль артикула в Excel';$issues[]='Дубль артикула в Excel';}
            elseif(count($matches)>1){$status='Дубль артикула на сайте';$issues[]='Дубль артикула на сайте: '.implode(', ',array_keys($matches));}
            elseif($guidMatch)$status='GUID найден у другого артикула на сайте';
            elseif($occupied)$status='GUID занят другим товаром на сайте';
            elseif(count($matches)===0)$status='Товар не найден на сайте';
            elseif($r['price_error'])$status='Проблема с ценой в Excel';
            else $status=$r['old_guid']===$r['guid']?'GUID уже совпадает':'Готов к обновлению';
            $blocking=$a===''||($counts[$a]??0)>1||count($matches)!==1||!empty($occupied)||!empty($issues);
            $r['status']=$status;$r['issues']=implode('; ',array_unique(array_merge($issues,$warnings)));
            $r['result']=$blocking?'Пропустить':'К обновлению';
        }
        unset($r);$job['base']=$base;$job['phase']='planned';
    }
    public static function update(&$r,$job,$cfg,$link){if(!empty($r['date_issue'])){$r['result']='Пропущено: артикул сохранён как дата';return;}$r['result']='Запись начата — проверить';[$index,$owners,$items,$base]=self::catalog($cfg,$link);$matches=$index[self::key($r['article'])]??[];if(count($matches)!==1||!isset($matches[$r['id']])){$r['result']='Пропущено: изменилось сопоставление';return;}$e=$matches[$r['id']];$occupied=$owners[self::key($r['guid'])]??[];unset($occupied[$r['id']]);if($occupied){$r['result']='Пропущено: GUID занят';return;}if(!isset($e['quantity'])||isset($e['price']['ambiguous'])){$r['result']='Пропущено: проблема каталога/цены';return;}if($r['price']===null){$r['result']='Пропущено: неверная цена';return;}
        $errors=[];$el=new \CIBlockElement();if(!$el->Update($r['id'],['XML_ID'=>$r['guid'],'ACTIVE'=>'Y']))$errors[]='Элемент: '.$el->LAST_ERROR;
        $actual=\CIBlockElement::GetByID($r['id'])->Fetch();if(!$actual||(string)$actual['XML_ID']!==$r['guid']||$actual['ACTIVE']!=='Y')$errors[]='XML_ID/ACTIVE не подтверждены';
        if(!$errors){$kod=self::kod($r['ib']);\CIBlockElement::SetPropertyValuesEx($r['id'],$r['ib'],[$kod['id']=>$kod['yes']]);$p=\CIBlockElement::GetProperty($r['ib'],$r['id'],[],['ID'=>$kod['id']])->Fetch();if(!$p||(int)$p['VALUE']!==$kod['yes'])$errors[]='KOD_IZ_1C не подтверждён';}
        if(!$errors){$q=\Bitrix\Catalog\Model\Product::update($r['id'],['QUANTITY'=>100]);if(!$q->isSuccess())$errors[]=implode('; ',$q->getErrorMessages());}
        if(!$errors){$existing=$e['price'];if($existing){$res=\Bitrix\Catalog\Model\Price::update($existing['id'],['PRICE'=>$r['price'],'CURRENCY'=>'RUB']);}else{$res=\Bitrix\Catalog\Model\Price::add(['PRODUCT_ID'=>$r['id'],'CATALOG_GROUP_ID'=>$base,'PRICE'=>$r['price'],'CURRENCY'=>'RUB']);}if(!$res->isSuccess())$errors[]='Цена: '.implode('; ',$res->getErrorMessages());}
        $after=\Bitrix\Catalog\ProductTable::getList(['select'=>['QUANTITY'],'filter'=>['=ID'=>$r['id']]])->fetch();$pr=\Bitrix\Catalog\PriceTable::getList(['select'=>['PRICE','CURRENCY'],'filter'=>['=PRODUCT_ID'=>$r['id'],'=CATALOG_GROUP_ID'=>$base]])->fetch();if(!$after||(float)$after['QUANTITY']!==100.0)$errors[]='Количество не подтверждено';if(!$pr||$pr['CURRENCY']!=='RUB'||abs((float)$pr['PRICE']-(float)$r['price'])>0.009)$errors[]='Цена не подтверждена';$r['result']=$errors?'Ошибка/частичная запись: '.implode('; ',array_unique($errors)):'Обновлено';
    }
}

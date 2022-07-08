<?php
include_once(dirname(__FILE__).'/install/demo.php');

if(!class_exists('CDATAExportExcelRunner'))
{
	class CDATAExportExcelRunner
	{
		protected static $moduleId = 'data.exportexcel';
		
		static function GetModuleId()
		{
			return self::$moduleId;
		}
		
		static function DemoExpired()
		{
			$DemoMode = CModule::IncludeModuleEx(self::$moduleId);
			$cnstPrefix = str_replace('.', '_', self::$moduleId);
			if ($DemoMode==MODULE_DEMO) {
				$now=time();
				if (defined($cnstPrefix."_OLDSITEEXPIREDATE")) {
					if ($now>=constant($cnstPrefix.'_OLDSITEEXPIREDATE') || constant($cnstPrefix.'_OLDSITEEXPIREDATE')>$now+2000000 || $now - filectime(__FILE__)>2000000) {
						return true;
					}
				} else{ 
					return true;
				}
			} elseif ($DemoMode==MODULE_DEMO_EXPIRED) {
				return true;
			}
			return false;
		}
		
		static function ExportIblock($params=array(), $fparams=array(), $stepparams=false, $pid = false)
		{
			if(self::DemoExpired()) return array();
			$ee = new CDATAExportExcel($params, $fparams, $stepparams, $pid);
			return $ee->Export();
		}
		
		static function ExportHighloadblock($params=array(), $fparams=array(), $stepparams=false, $pid = false)
		{
			if(self::DemoExpired()) return array();
			$ee = new CDATAExportExcelHighload($params, $fparams, $stepparams, $pid);
			return $ee->Export();
		}
	}
}

$moduleId = CDATAExportExcelRunner::GetModuleId();
$moduleJsId = str_replace('.', '_', $moduleId);
$pathJS = '/bitrix/js/'.$moduleId;
$pathCSS = '/bitrix/panel/'.$moduleId;
$pathLang = BX_ROOT.'/modules/'.$moduleId.'/lang/'.LANGUAGE_ID;
CModule::AddAutoloadClasses(
	$moduleId,
	array(
		'CDATAEEFieldList' => 'classes/general/field_list.php',
		'CDATAExportProfile' => 'classes/general/profile.php',
		'CDATAExportProfileAll' => 'classes/general/profile.php',
		'CDATAExportProfileDB' => 'classes/general/profile_db.php',
		'CDATAExportProfileFS' => 'classes/general/profile_fs.php',
		'CDATAExportExcel' => 'classes/general/export.php',
		'CDATAExportExcelStatic' => 'classes/general/export.php',
		'CDATAExportExcelHighload' => 'classes/general/export_highload.php',
		'CDATAExportExcelWriterXlsx' => 'classes/general/export_writer_xlsx.php',
		'CDATAExportExcelWriterCsv' => 'classes/general/export_writer_csv.php',
		'CDATAExportExcelWriterDbf' => 'classes/general/export_writer_dbf.php',
		'CDATAExportExtraSettings' => 'classes/general/extrasettings.php',
		'CDATAExportUtils' => 'classes/general/utils.php',
		//'CDATAExportCondTree' => 'classes/general/cond_tree.php',
		'\Bitrix\DataExportexcel\ProfileTable' => "lib/profile.php",
		'\Bitrix\DataExportexcel\ProfileHlTable' => "lib/profile_hl.php"
	)
);

$initFile = $_SERVER["DOCUMENT_ROOT"].BX_ROOT.'/php_interface/include/'.$moduleId.'/init.php';
if(file_exists($initFile)) include_once($initFile);

$arJSDataIBlockConfig = array(
	$moduleJsId => array(
		'js' => $pathJS.'/script.js',
		'css' => $pathCSS.'/styles.css',
		'rel' => array('jquery', $moduleJsId.'_chosen'/*, 'core_condtree'*/),
		'lang' => $pathLang.'/js_admin.php'
	),
	$moduleJsId.'_highload' => array(
		'js' => $pathJS.'/script_highload.js',
		'css' => $pathCSS.'/styles.css',
		'rel' => array('jquery', $moduleJsId.'_chosen'/*, 'core_condtree'*/),
		'lang' => $pathLang.'/js_admin_hlbl.php',
	),
	$moduleJsId.'_chosen' => array(
		'js' => $pathJS.'/chosen/chosen.jquery.min.js',
		'css' => $pathJS.'/chosen/chosen.min.css',
		'rel' => array('jquery')
	),
);

foreach ($arJSDataIBlockConfig as $ext => $arExt) {
	CJSCore::RegisterExt($ext, $arExt);
}
?>
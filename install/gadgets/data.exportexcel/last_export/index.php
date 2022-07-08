<?
use Bitrix\Main\Localization\Loc;
use Bitrix\Main\Loader;
Loc::loadMessages(__FILE__);
if (!defined("B_PROLOG_INCLUDED") || B_PROLOG_INCLUDED!==true)die();
$moduleId = 'data.exportexcel';
if(!Loader::includeModule($moduleId)) return;

$arGadgetParams["PROFILES_COUNT"] = (int)$arGadgetParams["PROFILES_COUNT"];
if ($arGadgetParams["PROFILES_COUNT"] <= 0)
	$arGadgetParams["PROFILES_COUNT"] = 10;

$oProfile = new \CDATAExportProfile();
$arProfiles = $oProfile->GetLastImportProfiles($arGadgetParams["PROFILES_COUNT"]);
if(!empty($arProfiles))
{
	echo '<table border="1">'.
		'<tr>'.
			'<th>'.Loc::getMessage('GD_DATA_EE_PROFILE_ID').'</th>'.
			'<th>'.Loc::getMessage('GD_DATA_EE_PROFILE_NAME').'</th>'.
			'<th>'.Loc::getMessage('GD_DATA_EE_PROFILE_DATE_START').'</th>'.
			'<th>'.Loc::getMessage('GD_DATA_EE_PROFILE_DATE_FINISH').'</th>'.
			'<th>'.Loc::getMessage('GD_DATA_EE_PROFILE_STATUS').'</th>'.
		'</tr>';
	foreach($arProfiles as $arProfile)
	{
		$arStatus = $oProfile->GetStatus($arProfile, true);
		echo '<tr'.($arStatus['STATUS']=='ERROR' ? ' style="background: #ffdddd;"' : '').'>'.
				'<td>'.$arProfile['ID'].'</td>'.
				'<td>'.$arProfile['NAME'].'</td>'.
				'<td>'.(is_callable(array($arProfile['DATE_START'], 'toString')) ? $arProfile['DATE_START']->toString() : '').'</td>'.
				'<td>'.(is_callable(array($arProfile['DATE_FINISH'], 'toString')) ? $arProfile['DATE_FINISH']->toString() : '').'</td>'.
				'<td>'.$arStatus['MESSAGE'].'</td>'.
			'</tr>';
	}
	echo '</table>';
}
else
{
	echo Loc::getMessage('GD_DATA_EE_NO_DATA');
}
?>



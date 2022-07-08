<?
require_once($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_admin_before.php");
require_once($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/iblock/prolog.php");
$moduleId = 'data.exportexcel';
$moduleEmail = 'app@kabubu.org';
$imgPath = '/bitrix/panel/'.$moduleId.'/images/video_icons/';
CModule::IncludeModule($moduleId);
IncludeModuleLangFile(__FILE__);
require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/prolog_popup_admin.php");
?>

<div class="data-ee-tabs" id="data-ee-help-tabs">
	<div class="data-ee-tabs-heads">
		<a href="javascript:void(0)" onclick="EHelper.SetTab(this);" class="active" title="<?echo htmlspecialcharsex(GetMessage("DATA_EE_HELP_TAB1_ALT"));?>"><?echo GetMessage("DATA_EE_HELP_TAB1");?></a>
		<a href="javascript:void(0)" onclick="EHelper.SetTab(this);" title="<?echo htmlspecialcharsex(GetMessage("DATA_EE_HELP_TAB2_ALT"));?>"><?echo GetMessage("DATA_EE_HELP_TAB2");?></a>
	</div>
	<div class="data-ee-tabs-bodies">
		<div class="active">
			<div>&nbsp;</div>
			<div class="data-ee-video-list">
				<a href="https://www.youtube.com/watch?v=QL2x2cyIPbI" target="_blank" title="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_COMMON"));?>">
					<img src="<?echo $imgPath;?>common.jpg" width="196px" height="110px" alt="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_COMMON"));?>" title="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_COMMON"));?>">
					<span><?echo GetMessage("DATA_EE_HELP_VIDEO_COMMON");?></span>
				</a>
				<a href="https://www.youtube.com/watch?v=IsBwPtRR2Po" target="_blank" title="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_PRICELIST"));?>">
					<img src="<?echo $imgPath;?>pricelist.jpg" width="196px" height="110px" alt="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_PRICELIST"));?>" title="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_PRICELIST"));?>">
					<span><?echo GetMessage("DATA_EE_HELP_VIDEO_PRICELIST");?></span>
				</a>
				<a href="https://www.youtube.com/watch?v=ha29fwucN1A" target="_blank" title="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_FILTERING"));?>">
					<img src="<?echo $imgPath;?>filtering.jpg" width="196px" height="110px" alt="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_FILTERING"));?>" title="<?echo htmlspecialcharsbx(GetMessage("DATA_EE_HELP_VIDEO_FILTERING"));?>">
					<span><?echo GetMessage("DATA_EE_HELP_VIDEO_FILTERING");?></span>
				</a>
			</div>
			<div>&nbsp;</div>
		</div>
		<div>
			<div>&nbsp;</div>
			<p class="data-ee-help-faq-prolog"><i><?echo sprintf(GetMessage("DATA_EE_FAQ_PROLOG"), $moduleEmail, $moduleEmail);?></i></p>
			<ol id="data-ee-help-faq">
				<!--<li>
					<a href="#"><?echo GetMessage("DATA_EE_FAQ_QUEST_PICTURES");?></a>
					<div><?echo GetMessage("DATA_EE_FAQ_ANS_PICTURES");?></div>
				</li>-->
			</ol>
			<div>&nbsp;</div>
		</div>
	</div>
</div>
<?require($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include/epilog_popup_admin.php");?>
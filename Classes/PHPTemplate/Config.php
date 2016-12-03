<?php
/* PHPTemplate_Config
 *
* @category   PHPTemplate
* @package    PHPTemplate_Config
* @copyright
*/
class PHPTemplate_Config
{
	/**
	 *  シート設定リスト
	 *
	 *  @var  array(PHPTemplate_SheetConfig)
	 */
	private $_sheetConfigArray = array();


	/**
	 * シート数取得
	 *
	 * @return int
	 */
	public function getSheetConfigArrayCount(){
		return count($this->_sheetConfigArray);
	}

	/**
	 * シート設定リスト取得
	 * @return PHPTemplate_SheetConfig
	 */
	public function getSheetConfig($idx){

		//if($idx > count($this->_sheetConfigArray)){
		//	throw new PHPTemplate_Exception('sheet index '.($idx).' config is not found.');
		//}
		//return $this->_sheetConfigArray[$idx];

		if(array_key_exists($idx, $this->_sheetConfigArray)){
			return $this->_sheetConfigArray[$idx];
		}
		else{
			return NULL;
		}
	}

	/**
	 * XMLファイルを読み込む
	 * @param string $path
	 */
	public function readXML($path){
		$xml = @simplexml_load_file($path);
		if ($xml) {
			foreach ($xml->sheet as $s) {
				$sheetConfig = new PHPTemplate_SheetConfig();
				$sheetConfig->setSheetSetting($s);

				$this->_sheetConfigArray[$sheetConfig->getSheetIndex()] = $sheetConfig;
			}
		}
		else{
			throw new PHPTemplate_Exception('xml config file is invalid');
		}
	}
}
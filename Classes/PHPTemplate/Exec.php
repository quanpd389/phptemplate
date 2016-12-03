<?php
/**
 * PHPTemplate_Exec
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Exec
 * @copyright
 */
class PHPTemplate_Exec
{
	/**
	 *  スクリプト文字列
	 *
	 *  @var  string
	 */
	private $_script;

	/**
	 *	Create a new Exec
	 *
	 *	@param	string 	$text	exec parameter
	 */
	public function __construct($text)
	{
		$pos = strpos($text, "#exec");
		if($pos !== false){
			$this->_script = trim(substr($text, strlen("#exec")));
		}
		else{
			throw new PHPTemplate_Exception('#exec parameter is invalid');
		}
	}

	/**
	 * 制御ブロックの展開
	 *
	 * @param array				$data
	 * @param int				$rownum
	 * @param PHPTemplate_Page	$page
	 * @param bool				$dataF
	 */
	public function expand(&$data, &$rownum, &$page, $dataF=TRUE){

		$pos = strpos($this->_script, "=");
		if($pos !== false){
			$var = trim(substr($this->_script, 0, $pos));
			$scriptStr = trim(substr($this->_script, $pos+1));
			$varArray = Array();
			$tok = strtok($scriptStr, " \t()+-*%/|&=!<>;:?");
			while ($tok !== false) {
				if(!is_numeric($tok)){
					$varArray[] = $tok;
				}
				$tok = strtok(" \t()+-*%/|&=!<>;:?");
			}
			usort($varArray, "PHPTemplate_Util::cmp");
			// 変数からバインドする値を取得
			$cmd = $scriptStr;
			foreach ($varArray as $val){
				$p = PHPTemplate_Util::getBindData($val, $data, $rownum, $page->getPagenum(), FALSE);
				if(!is_null($p)){
					if(is_string($p)){
						$cmd = str_replace($val, "\"".$p."\"", $cmd);
					}
					else{
						$cmd = str_replace($val, $p, $cmd);
					}
				}
			}

			$text="return ". $cmd. ";";
			$ret=eval($text);
			$data[$var] = $ret;
			return;
		}
	}
}
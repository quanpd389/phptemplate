<?php
/**
 * PHPTemplate_Util
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Util
 * @copyright
 */
class PHPTemplate_Util
{
	/* Row types */
	const ROW_TYPE_NONE      = 0;	//
	const ROW_TYPE_FOREACH   = 1;	//
	const ROW_TYPE_HFOREACH  = 2;	//
	const ROW_TYPE_WHILE     = 3;	//
	const ROW_TYPE_IF        = 4;	//
	const ROW_TYPE_ELSEIF    = 5;	//
	const ROW_TYPE_ELSE      = 6;	//
	const ROW_TYPE_END       = 7;	//
	const ROW_TYPE_COMMENT   = 8;	//
	const ROW_TYPE_VAR       = 9;	//
	const ROW_TYPE_EXEC      = 10;	//
	const ROW_TYPE_PAGEBREAK = 11;  //
	const ROW_TYPE_PAGEHEADER= 12;	//
	const ROW_TYPE_PAGEFOOTER= 13;	//
	const ROW_TYPE_SUSPEND   = 14;	//
	const ROW_TYPE_RESUME    = 15;	//
	const ROW_TYPE_LINK_URL  = 16;	//
	const ROW_TYPE_LINK_EMAIL= 17;	//
	const ROW_TYPE_LINK_FILE = 18;	//
	const ROW_TYPE_LINK_THIS = 19;	//
	const ROW_TYPE_IMG       = 20;	//
	const ROW_TYPE_BLOCK     = 21;	//
	const ROW_TYPE_FONT      = 22;	//

	/* default row size*/
	const DEFAULT_ROW_SIZE		= 13.50;
	/* default column size*/
	const DEFAULT_COLUMN_SIZE	= 8.38;

	/* import filter type */
	const FILTER_TYPE_TEXT       = 0;	// 文字データ
	const FILTER_TYPE_FONT_COLOR = 1;	// 文字色
	const FILTER_TYPE_CELL_COLOR = 2;	// セル色

	/* excel type */
	const EXCEL_TYPE_95			= 'Excel5';		// Excel95
	const EXCEL_TYPE_2007		= 'Excel2007';	// Excel2007

	/**
	 * 制御文か判定　制御文なら種別を返す
	 *
	 * @param	string		$text
	*/
	static public function getControlKind($text, $column) {
		if(is_null($text))
			return self::ROW_TYPE_NONE;
		if(strlen($text) == 0)
			return self::ROW_TYPE_NONE;

		$t = strtok(trim($text), " \t(");
		if($t == "#foreach" && $column=='A'){
			return self::ROW_TYPE_FOREACH;
		}
		elseif($t == "#hforeach"){
			return self::ROW_TYPE_HFOREACH;
		}
		elseif($t == "#while" && $column=='A'){
			return self::ROW_TYPE_WHILE;
		}
		elseif($t == "#if" && $column=='A'){
			return self::ROW_TYPE_IF;
		}
		elseif($t == "#else" && $column=='A'){
			if(strtok(" \t") == "if")
				return self::ROW_TYPE_ELSEIF;
			else
				return self::ROW_TYPE_ELSE;
		}
		elseif($t == "#comment"){
			return self::ROW_TYPE_COMMENT;
		}
		elseif($t == "#var" && $column=='A'){
			return self::ROW_TYPE_VAR;
		}
		elseif($t == "#exec" && $column=='A'){
			return self::ROW_TYPE_EXEC;
		}
		elseif ($t == "#pageBreak" && $column=='A'){
			return self::ROW_TYPE_PAGEBREAK;
		}
		elseif ($t == "#pageHeaderStart" && $column=='A'){
			return self::ROW_TYPE_PAGEHEADER;
		}
		elseif ($t == "#pageFooterStart" && $column=='A'){
			return self::ROW_TYPE_PAGEFOOTER;
		}
		elseif($t == "#end"){
			return self::ROW_TYPE_END;
		}
		elseif($t == "#link-url"){
			return self::ROW_TYPE_LINK_URL;
		}
		elseif($t == "#link-email"){
			return self::ROW_TYPE_LINK_EMAIL;
		}
		elseif($t == "#link-file"){
			return self::ROW_TYPE_LINK_FILE;
		}
		elseif($t == "#link-this"){
			return self::ROW_TYPE_LINK_THIS;
		}
		elseif ($t == "#img"){
			return self::ROW_TYPE_IMG;
		}
		elseif ($t == "#suspend"){
			return self::ROW_TYPE_SUSPEND;
		}
		elseif ($t == "#resume"){
			return self::ROW_TYPE_RESUME;
		}
		elseif ($t == "#font" && $column=='A'){
			return self::ROW_TYPE_FONT;
		}
		else{
			return self::ROW_TYPE_NONE;
		}

	}

	/**
	 * データリストから指定のキーを持つデータを取得する
	 *
	 * @param	string					$keyName
	 * @param	array					$dataList
	 * @param	int						$rownum
	 * @param	PHPTemplate_Page		$page
	 * @param	bool					$errF
	 * @return	string
	*/
	static public function getBindData($keyName, $dataList, $rownum, $page, $errF=TRUE) {
		$data = NULL;
		$key = $keyName;

		//keyに!が含まれるか
		$flg = FALSE;
		$def = NULL;
		$str = NULL;
		if(strstr($key, "!")!=NULL){
			$flg = TRUE;
			$def = substr(strstr($key, "!"), 1);
			$str = strstr($key, "!", TRUE);
		}
		else{
			$str = $key;
		}

		//内部変数rownumか？
		if($key == "rownum"){
			$data = (string)$rownum;
			return $data;
		}

		// オブジェクト参照型か？
		$name = NULL;
		if(strstr($str, ".")!=NULL){
			$name = substr(strstr($str, "."), 1);
			$str = strstr($str, ".", TRUE);
		}
		// Array参照型か？
		elseif(strstr($str, "[")!=NULL){
			$name = substr(strstr($str, "["), 1);
			$str = strstr($str, "[", TRUE);
		}
		//内部変数page.pagenumか？
		if ($str == "page" && $name == "pagenum"){
			$data = (string)$page;
			return $data;
		}

		if(array_key_exists($str, $dataList)){
			if($name){//オブジェクト参照型
				if(is_object($dataList[$str])){
					//$nameをドットで区切る
					$t = strtok($name, " .");
					$obj=$dataList[$str];
					while ($t !== false) {
						$name = strtoupper(substr($t, 0, 1)).substr($t, 1);
						$name = "get".$name;
						if(method_exists($obj,$name)){
							$obj = call_user_func(array($obj, $name));
						}
						else{
							$data = PHPTemplate_Util::getBindErrorData($flg, $def, $errF);
							return $data;
						}
						$t = strtok(".");
					}
					$data = $obj;
				}
				elseif(is_array($dataList[$str])){
					//$nameを[で区切る
					$t = strtok($name, "[");
					$array = $dataList[$str];
					while ($t !== false) {
						$key = strstr($name, "]", TRUE);
						$key = trim($key, "'\"");
						$array = $array[$key];
						//echo($array."\n");
						$t = strtok("[");
					}
					$data = $array;
				}
				else{
					//データなし
					$data = PHPTemplate_Util::getBindErrorData($flg, $def, $errF);
				}
			}else{
				$data = $dataList[$str];
			}
		}else{
			//データなし
			$data = PHPTemplate_Util::getBindErrorData($flg, $def, $errF);
		}
		return $data;
	}

	/**
	 * バインドエラー時のデータ取得
	 *
	 * @param		bool		$flg
	 * @param		string		$def
	 * @param		bool		$errF
	*/
	static private function getBindErrorData($flg, $def, $errF){
		if(!$errF)	return NULL;

		//!なし　エラー
		if(!$flg){
			return "エラー！";
		}
		//!あり　空白orデフォルト値
		else{
			if(!$def){
				return "";
			}
			else{
				return $def;
			}
		}

	}

	/**
	 * 制御文字列から指定のパラメータ値を取りだす
	 * ex.index=idx
	 *
	 * @param	string		$text		制御文字列
	 * @param	string		$param		パラメータ名
	*/
	static public function getParamater($text, $param){

		if(strlen($text) > 0){
			$pos = strpos($text, $param);
			if($pos !== false){
				$paramText = trim(substr($text, $pos+strlen($param)));//= idx xxxxx
				if(substr($paramText, 0, 1) == "="){
					$paramText = trim(substr($paramText, 1));
					$pos = strpos($paramText, " ");
					if($pos === false){
						return  $paramText;
					}
					else{
						return trim(substr($paramText, 0, $pos));
					}
				}
			}
		}
		return NULL;
	}

	/**
	 * 条件文の評価
	 *
	 * @param	string		$text
	 * @param	array		$param
	 * @param	int			$rownum
	 * @param	int			$pagenum
	 * @return	bool
	 */
	static public function evalCondition($text, $param, $rownum, $pagenum){
		if($text == NULL)	return TRUE;

		// 条件文から変数部分を抜き出す index,list.var1, list.size()...
		$varArray = Array();
		$tok = strtok($text, " \t()+-*%/|&=!<>;:?,");
		while ($tok !== false) {
			if(!is_numeric($tok)){
				$varArray[] = $tok;
			}
			$tok = strtok(" \t()+-*%/|&=!<>;:?,");
		}
		usort($varArray, "PHPTemplate_Util::cmp");
		// 変数からバインドする値を取得
		$cmd = $text;
		$varText = "";
		foreach ($varArray as $val){
//echo("val:".$val."\n");
			$p = PHPTemplate_Util::getBindData($val, $param, $rownum, $pagenum, FALSE);
			if(!is_null($p)){
				if(is_string($p)){
					$p = str_replace("\"", "\\\"", $p);
					$cmd = str_replace($val, "\"".$p."\"", $cmd);
				}
				elseif (is_array($p)){
					$varText = "$".$val."=".var_export($p,true).";";
					$cmd = str_replace($val, "$".$val, $cmd);
				}
				else{
					$cmd = str_replace($val, $p, $cmd);
				}
			}
		}
		$text=$varText."if( ". $cmd. "){return TRUE;} else{return FALSE;}";
//echo("eval:".$text."\n");
		$ret=eval($text);
		return $ret;
	}
	/**
	 * 文字数の長い順に並べ帰るための比較関数
	 *
	 * @param	string		$a
	 * @param	string		$b
	 */
	static public function cmp($a, $b)
	{
		if (strlen($a) == strlen($b)) {
			return 0;
		}
		return (strlen($a) > strlen($b)) ? -1 : 1;
	}


}
<?php
/**
 * PHPTemplate_Row
 *
 * @category   PHPTemplate
 * @package    PHPTemplate_Row
 * @copyright
 */

class PHPTemplate_Row
{
	/**
	 * 行内のセルリスト
	 *  @var  array
	 */
	private $_cellArray = array();

	/**
	 * 初期状態での行番号
	 * @var  int
	 */
	private $_baseRow = -1;

	/**
	 * コピーして追加した行数
	 * @var  int
	 */
	private $_addRowCnt = 0;

	/**
	 * 行の結合情報
	 * @var  PHPTemplate_RowMerge[]
	 */
	private $_rowMergeInfo = array();

	/**
	 *	Create a new Row
	 *
	 *	@param	PHPExcel_WorkSheet	$sheet
	 */
	public function __construct()
	{
	}

	/**
	 * コピーして追加した行数の取得
	 */
	public function getAddRowCnt() {
		return $this->_addRowCnt;
	}

	/**
	 * 制御行（コピー元の行）の削除
	 */
	public function delRow(&$cnt){
		if($this->_addRowCnt>0 && count($this->_cellArray) > 0){
			$sheet = $this->_cellArray[0]->getWorksheet();
			$sheet->removeRow($this->_baseRow + $cnt);
			$cnt--;
			$this->_addRowCnt--;
		}
	}

	/**
	 * セル追加
	 * @param		PHPTemplate_Cell		$objCell
	 */
	public function addCell($objCell) {

		//$objCell = new PHPTemplate_Cell($cell);

		$this->_cellArray[] = $objCell;

		$this->_baseRow = $objCell->getRow();

		//行の結合情報をセット
		if($objCell->getMergeRowStart()){
			$merge = new PHPTemplate_RowMerge($objCell->getMergeRowCnt(), $objCell->getMergeInfo(), $objCell->getStyle());
			$this->_rowMergeInfo[] = $merge;
		}

	}

	/**
	 * セル取得
	 * @param		int		$i		インデックス
	 */
	public function getCell($i) {
		return $this->_cellArray[$i];
	}

	/**
	 * セル数取得
	 * @return		int
	 */
	public function getCellCnt() {
		return count($this->_cellArray);

	}

	/**
	 * 制御列追加
	 *
	 * @param	mixed	$obj	制御オブジェクト
	 */
	public function addControl($obj) {

		$this->_cellArray[] = $obj;

	}

	/**
	 * 行の結合情報取得
	 *
	 * @return	array
	 */
	public function getRowMergeInfo()
	{
		return $this->_rowMergeInfo;
	}

	/**
	 * 行を展開
	 *
	 * @param array				$data
	 * @param int				$rownum
	 * @param PHPTemplate_Page	$page
	 */
	public function expand(&$data, &$rownum, &$page) {
		$mergeInfo = Array();
		$sheet = NULL;
		$cnt =0;
		foreach ($this->_cellArray as $cell) {
			$sheet = $cell->getWorksheet();
			if($cnt==0){
				//行高さ設定
				$sheet->getRowDimension($rownum)->setRowHeight($cell->getRowHeight());
			}
			$text = $cell->getText();
			if (PHPTemplate_Util::getControlKind($text, $cell->getColumn())==PHPTemplate_Util::ROW_TYPE_SUSPEND){
				$val = trim(substr($text, strlen("#suspend")));
				$key = trim($val,'\$\{\}');
				$cell->setRow($rownum);
				$page->addSuspend($key, $cell);
			}
			$cell->bindData($data, $rownum, $page);

			//セルの結合情報
			$info = $cell->getMergeInfo();
			if($info!=NULL){
				if($cell->getMergeRowStart() /*&& $cell->getMergeRowCnt() > 0*/){
					$merge = $info[0].(string)($rownum).":".$info[1].(string)($rownum + $cell->getMergeRowCnt());
					$sheet->mergeCells($merge);
					//echo "Row".$merge;
				}
			}
			$cnt++;
		}

		$rownum++;

	}

	/**
	 * 行を追加
	 *
	 * @param int				$rownum
	 * @param bool				$cpyValue		true:値をコピーする
	 */
	public function addRow($rowNum, $cpyValue=TRUE){
		$cnt = 0;
		$newRow = new PHPTemplate_Row();
		$sheet = NULL;
		$mergeInfo = Array();
		foreach ($this->_cellArray as $objCell){
			$sheet = $objCell->getWorksheet();
			if($cnt==0){
				//$sheet->insertNewRowBefore($rowNum);
				//行高さ設定
				$sheet->getRowDimension($rowNum)->setRowHeight($objCell->getRowHeight());
				$this->_addRowCnt++;
			}
			//スタイルコピー
			$style = $objCell->getStyle();
			$sheet->duplicateStyle(clone $style, $objCell->getColumn().($rowNum));
			if($cpyValue){
				$sheet->setCellValue($objCell->getColumn().($rowNum),$objCell->getText());
			}
			$cell = $sheet->getCellByColumnAndRow($cnt, $rowNum);
			$info = $objCell->getMergeInfo();
			$newRow->addCell(new PHPTemplate_Cell($cell, $objCell->getMergeRowHeight(), $objCell->getMergeColumnWidth(), $objCell->getMergeRowStart(), $objCell->getMergeRowCnt(), $info));

			$cnt = $cnt+1;

			//セルの結合情報

			//echo "rownum:".$rowNum."\n";
			//echo "margestart:".$objCell->getMergeRowStart()."\n";
			//echo "rowcount:".$objCell->getMergeRowCnt()."\n";
			//print_r($info);


			if($info!=NULL){
				if($objCell->getMergeRowStart() && $objCell->getMergeRowCnt() == 0){
					$addF = TRUE;
					foreach($mergeInfo as $m){
						if($m[0]==$info[0] && $m[1]==$info[1]){
							$addF = FALSE;
							break;
						}
					}
					if($addF){
						$mergeInfo[] = $info;
					}
				}
			}
		}
		/*foreach($mergeInfo as $m){
			$merge = $m[0].(string)($rowNum).":".$m[1].(string)($rowNum);
			$sheet->mergeCells($merge);
			echo "Row".$merge;
		}*/

		/*foreach ($this->_rowMergeInfo as $info){
			$merge = new PHPTemplate_RowMerge($info->getRowCnt(), $info->getMergeInfo(), $info->getStyle());
			$newRow->_rowMergeInfo[] = $merge;
		}*/


		return $newRow;

	}

	/**
	 * 空白セルを詰める
	 * @param int				$cnt
	 */
	public function trimRow($cnt){
		while(count($this->_cellArray) > $cnt){
			array_pop($this->_cellArray);
		}
	}


}
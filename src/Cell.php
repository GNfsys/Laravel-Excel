<?php

namespace Maatwebsite\Excel;

use PhpOffice\PhpSpreadsheet\Calculation\Exception;
use PhpOffice\PhpSpreadsheet\Cell\Cell as SpreadsheetCell;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/** @mixin SpreadsheetCell */
class Cell
{
    use DelegatedMacroable;

    /**
     * @var SpreadsheetCell
     */
    private $cell;

    /**
     * @param  SpreadsheetCell  $cell
     */
    public function __construct(SpreadsheetCell $cell)
    {
        $this->cell = $cell;
    }

    /**
     * @param  Worksheet  $worksheet
     * @param  string  $coordinate
     * @return Cell
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function make(Worksheet $worksheet, string $coordinate)
    {
        return new static($worksheet->getCell($coordinate));
    }

    /**
     * @return SpreadsheetCell
     */
    public function getDelegate(): SpreadsheetCell
    {
        return $this->cell;
    }

    /**
     * @param  null  $nullValue
     * @param  bool  $calculateFormulas
     * @param  bool  $formatData
     * @return mixed
     */
    public function getValue($nullValue = null, $calculateFormulas = false, $formatData = true)
    {
        return self::getValueFromSpreadsheetCell($this->cell, $nullValue, $calculateFormulas, $formatData);
    }

    public static function getValueFromSpreadsheetCell(SpreadsheetCell $cell, $nullValue = null, $calculateFormulas = false, $formatData = true)
    {
        $value = $nullValue;

        $cellValue = $cell->getValue();

        if ($cellValue !== null) {
            if ($cellValue instanceof RichText) {
                $value = $cellValue->getPlainText();
            } elseif ($calculateFormulas) {
                try {
                    $value = $cell->getCalculatedValue();
                } catch (Exception) {
                    $value = $cell->getOldCalculatedValue();
                }
            } else {
                $value = $cellValue;
            }

            if ($formatData) {
                $style = $cell->getWorksheet()->getParent()->getCellXfByIndex($cell->getXfIndex());
                $value = NumberFormat::toFormattedString(
                    $value,
                    ($style && $style->getNumberFormat()) ? $style->getNumberFormat()->getFormatCode() : NumberFormat::FORMAT_GENERAL
                );
            }
        }

        foreach (config('excel.imports.cells.middleware', []) as $pipe) {
            $value = $pipe($value);
        }

        return $value;
    }
}

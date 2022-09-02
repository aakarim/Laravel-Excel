<?php

namespace Maatwebsite\Excel\Events;

use Maatwebsite\Excel\PhpSpreadsheetSheet;

class BeforeSheet extends Event
{
    /**
     * @var Sheet
     */
    public $sheet;

    /**
     * @var object
     */
    private $exportable;

    /**
     * @param  Sheet  $sheet
     * @param  object  $exportable
     */
    public function __construct(PhpSpreadsheetSheet $sheet, $exportable)
    {
        $this->sheet       = $sheet;
        $this->exportable  = $exportable;
    }

    /**
     * @return PhpSpreadsheetSheet
     */
    public function getSheet(): PhpSpreadsheetSheet
    {
        return $this->sheet;
    }

    /**
     * @return object
     */
    public function getConcernable()
    {
        return $this->exportable;
    }

    /**
     * @return mixed
     */
    public function getDelegate()
    {
        return $this->sheet;
    }
}

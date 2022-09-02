<?php

namespace Maatwebsite\Excel\Events;

use Maatwebsite\Excel\PhpSpreadsheetWriter;

class AfterReopen extends Event
{
    /**
     * @var Writer
     */
    public $writer;

    /**
     * @var object
     */
    private $exportable;

    /**
     * @param  Writer  $writer
     * @param  object  $exportable
     */
    public function __construct(PhpSpreadsheetWriter $writer, $exportable)
    {
        $this->writer     = $writer;
        $this->exportable = $exportable;
    }

    /**
     * @return PhpSpreadsheetWriter
     */
    public function getWriter(): PhpSpreadsheetWriter
    {
        return $this->writer;
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
        return $this->writer;
    }
}

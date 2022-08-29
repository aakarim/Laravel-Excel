<?php

namespace Maatwebsite\Excel\Events;

use Maatwebsite\Excel\Writer;

class AfterWriting extends Event
{
    /**
     * @var Writer
     */
    public $writer;

    /**
     * @var string
     */
    private $path;

    /**
     * @param  Writer  $writer
     * @param  string  $path
     */
    public function __construct(Writer $writer, string $path)
    {
        $this->writer     = $writer;
        $this->path = $path;
    }

    /**
     * @return Writer
     */
    public function getWriter(): Writer
    {
        return $this->writer;
    }

    /**
     * @return object
     */
    public function getConcernable()
    {
        return $this->path;
    }

    /**
     * @return mixed
     */
    public function getDelegate()
    {
        return $this->writer;
    }
}

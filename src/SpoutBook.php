<?php
namespace Maatwebsite\Excel;

class SpoutBook
{
    public array $spreadsheet;
    
    public function __construct(array $spreadsheeet)
    {
        $this->spreadsheet = $spreadsheeet;   
    }
}
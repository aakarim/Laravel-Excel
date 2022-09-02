<?php
namespace Maatwebsite\Excel\Contracts;

use Maatwebsite\Excel\Exporter;

interface Sheet {
    public function appendRows(iterable $rows, mixed $sheetExport);
}
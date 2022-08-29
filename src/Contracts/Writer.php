<?php
namespace Maatwebsite\Excel\Contracts;

use Maatwebsite\Excel\Files\TemporaryFile;

interface Writer {
    public function reopen(TemporaryFile $tmpFile, string $writerType): Writer;

    public function getSheetByIndex(int $index): Sheet;

    public function write(mixed $exportable, TemporaryFile $tmpFile, string $writeType);
}
<?php

namespace Maatwebsite\Excel;

use Box\Spout\Reader\ReaderInterface;
use Error;
use Illuminate\Contracts\Support\Arrayable;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\WithCustomStartCell;
use Maatwebsite\Excel\Concerns\WithCustomValueBinder;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Contracts\Sheet;
use Maatwebsite\Excel\Helpers\ArrayHelper;
use Maatwebsite\Excel\Helpers\CellHelper;

class SpoutSheet implements Sheet
{
    private array $worksheet;

    public function __construct(array $worksheet)
    {
        $this->worksheet = $worksheet;
    }

    public function appendRows(iterable $rows, Exporter $sheetExport)
    {
        if (method_exists($sheetExport, 'prepareRows')) {
            $rows = $sheetExport->prepareRows($rows);
        }

       $rows = (new Collection($rows))->flatMap(function ($row) use ($sheetExport) {
            if ($sheetExport instanceof WithMapping) {
                $row = $sheetExport->map($row);
            }

            if ($sheetExport instanceof WithCustomValueBinder) {
                throw new Error("spout does not support custom value bindings");
            }

            return ArrayHelper::ensureMultipleRows(
                static::mapArraybleRow($row)
            );
        })->toArray();

        array_push($this->worksheet, ...$rows);
    }

    /**
     * @param  mixed  $row
     * @return array
     */
    public static function mapArraybleRow($row): array
    {
        // When dealing with eloquent models, we'll skip the relations
        // as we won't be able to display them anyway.
        if (is_object($row) && method_exists($row, 'attributesToArray')) {
            return $row->attributesToArray();
        }

        if ($row instanceof Arrayable) {
            return $row->toArray();
        }

        // Convert StdObjects to arrays
        if (is_object($row)) {
            return json_decode(json_encode($row), true);
        }

        return $row;
    }
}
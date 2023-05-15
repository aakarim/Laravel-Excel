<?php

namespace Maatwebsite\Excel;

use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Carbon\Carbon;
use Error;
use Illuminate\Contracts\Support\Arrayable;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\WithCustomValueBinder;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Contracts\Sheet;
use Maatwebsite\Excel\Helpers\ArrayHelper;

class SpoutSheet implements Sheet
{
    private SpoutBook $worksheet;
    private int $index;

    public function __construct(SpoutBook $worksheet, int $index)
    {
        $this->worksheet = $worksheet;
        $this->index = $index;
    }

    public function appendRows(iterable $rows, $sheetExport)
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
        $rows = array_map(function($arr) {
            $style = (new StyleBuilder())
                ->setShouldWrapText(false)
                ->build();
            return WriterEntityFactory::createRowFromArray($arr, $style);
        }, $rows);
        if (is_null($this->worksheet->spreadsheet[$this->index] ?? null)) {
            $this->worksheet->spreadsheet[$this->index] = [];
        }
        array_push($this->worksheet->spreadsheet[$this->index], ...$rows);
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

        // TODO: move to a separate class for value binding
        $row = array_map(function ($r) {
            if ($r instanceof Carbon) {
                return $r->format('Y-m-d H:i:s');
            }
            if (is_string($r)) {
                // TODO: report
                if (strlen($r) > 32767) {
                    $r = substr($r, 0, 32767);
                }
            }
            return $r;
        }, $row);

        return $row;
    }
}
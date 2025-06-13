<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Illuminate\Support\Facades\Log;

class FieldExcelController extends Controller
{
    public function getFields($fileIndex, $sheetIndex)
    {
        $excelFiles = session('excel_files', []);

        if (!isset($excelFiles[$fileIndex])) {
            return redirect()->route('file.index')->with('error', 'File không tồn tại trong danh sách.');
        }

        $filePath = $excelFiles[$fileIndex]['path'];
        if (!file_exists($filePath)) {
            return redirect()->route('file.index')->with('error', 'File không tồn tại tại: ' . $filePath);
        }

        try {
            $spreadsheet = IOFactory::load($filePath);
            $sheetNames = $spreadsheet->getSheetNames();

            if (!isset($sheetNames[$sheetIndex])) {
                return redirect()->route('file.index')->with('error', 'Sheet không tồn tại.');
            }

            $worksheet = $spreadsheet->getSheet($sheetIndex);
            $highestColumn = $worksheet->getHighestColumn();
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

            $fields = [];
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $cell = $worksheet->getCellByColumnAndRow($col, 1);
                $value = $cell->getCalculatedValue();
                if ($value !== null && trim($value) !== '') {
                    $fields[] = $value;
                }
            }

            $sheetFields = session('sheet_fields', []);
            $sheetFields[$fileIndex][$sheetIndex] = [
                'sheet_name' => $sheetNames[$sheetIndex],
                'fields' => $fields,
            ];
            session(['sheet_fields' => $sheetFields]);

            return redirect()->route('file.index')->with('success', 'Đã lấy danh sách trường của sheet "' . $sheetNames[$sheetIndex] . '" thành công.');

        } catch (\Exception $e) {
            Log::error('Lỗi khi đọc trường: ' . $e->getMessage());
            return redirect()->route('file.index')->with('error', 'Không thể đọc trường: ' . $e->getMessage());
        }
    }

    public function removeFields($fileIndex, $sheetIndex)
    {
        $sheetFields = session('sheet_fields', []);
        $mappings = session('mappings', []);

        if (isset($sheetFields[$fileIndex][$sheetIndex])) {
            unset($sheetFields[$fileIndex][$sheetIndex]);
            if (empty($sheetFields[$fileIndex])) {
                unset($sheetFields[$fileIndex]);
            }
            // Xóa các mapping liên quan
            $mappings = array_filter($mappings, fn($mapping) => 
                $mapping['field']['file_index'] != $fileIndex || 
                $mapping['field']['sheet_index'] != $sheetIndex
            );
            $mappings = array_values($mappings);

            session([
                'sheet_fields' => $sheetFields,
                'mappings' => $mappings
            ]);

            return redirect()->route('file.index')->with('success', 'Đã xóa danh sách trường của sheet.');
        }

        return redirect()->route('file.index')->with('error', 'Danh sách trường không tồn tại.');
    }
}
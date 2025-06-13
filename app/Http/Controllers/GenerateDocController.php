<?php
namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpWord\TemplateProcessor;
use Illuminate\Support\Facades\Log;
use PhpOffice\PhpSpreadsheet\IOFactory;

class GenerateDocController extends Controller
{
    public function setPrimaryKey(Request $request)
    {
        $request->validate([
            'doc_index' => 'required|integer',
            'variable' => 'required|string',
        ]);

        $docIndex = $request->input('doc_index');
        $variable = $request->input('variable');
        $docVariables = session('doc_variables', []);

        if (isset($docVariables[$docIndex])) {
            $docVariables[$docIndex]['primary_key'] = $variable;
            session(['doc_variables' => $docVariables]);
            return redirect()->route('file.index')->with('success', 'Đã đặt khóa chính cho biến: ' . $variable);
        }

        return redirect()->route('file.index')->with('error', 'Không tìm thấy tài liệu.');
    }

    public function setOutputFolder(Request $request)
    {
        $request->validate([
            'output_folder' => 'required|string',
        ]);

        $outputFolder = trim($request->input('output_folder'), '"\'');
        $outputFolder = str_replace('/', '\\', $outputFolder);

        if (!is_dir($outputFolder)) {
            return redirect()->route('file.index')->with('error', 'Thư mục đầu ra không tồn tại: ' . $outputFolder);
        }

        session(['output_folder' => $outputFolder]);
        return redirect()->route('file.index')->with('success', 'Đã đặt thư mục đầu ra: ' . $outputFolder);
    }

    public function generateDoc($docIndex)
    {
        $docFiles = session('doc_files', []);
        $docVariables = session('doc_variables', []);
        $mappings = session('mappings', []);
        $excelFiles = session('excel_files', []);
        $outputFolder = session('output_folder');

        if (!isset($docFiles[$docIndex]) || !isset($docVariables[$docIndex])) {
            return redirect()->route('file.index')->with('error', 'Tài liệu không tồn tại.');
        }

        if (!$outputFolder || !is_dir($outputFolder)) {
            return redirect()->route('file.index')->with('error', 'Thư mục đầu ra chưa được thiết lập hoặc không tồn tại.');
        }

        $docPath = $docFiles[$docIndex]['path'];
        if (!file_exists($docPath)) {
            return redirect()->route('file.index')->with('error', 'File tài liệu không tồn tại tại: ' . $docPath);
        }

        try {
            $primaryKey = $docVariables[$docIndex]['primary_key'] ?? null;
            $relevantMappings = array_filter($mappings, fn($m) => $m['doc_index'] == $docIndex);

            if (empty($relevantMappings)) {
                return redirect()->route('file.index')->with('error', 'Không có mapping nào cho tài liệu này.');
            }

            $fileIndex = null;
            $sheetIndex = null;
            foreach ($relevantMappings as $mapping) {
                $fileIndex = $mapping['field']['file_index'];
                $sheetIndex = $mapping['field']['sheet_index'];
                break;
            }

            if (!isset($excelFiles[$fileIndex]) || !file_exists($excelFiles[$fileIndex]['path'])) {
                return redirect()->route('file.index')->with('error', 'File Excel không tồn tại.');
            }

            $spreadsheet = IOFactory::load($excelFiles[$fileIndex]['path']);
            $worksheet = $spreadsheet->getSheet($sheetIndex);
            $highestRow = $worksheet->getHighestRow();
            $highestColumn = $worksheet->getHighestColumn();
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

            $headerRow = [];
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $value = $worksheet->getCellByColumnAndRow($col, 1)->getCalculatedValue();
                $headerRow[$col] = $value ? trim($value) : '';
            }

            $primaryKeyCol = null;
            if ($primaryKey) {
                foreach ($relevantMappings as $mapping) {
                    if ($mapping['variable'] === $primaryKey) {
                        $fieldName = $mapping['field']['field'];
                        $primaryKeyCol = array_search($fieldName, $headerRow);
                        break;
                    }
                }
            }

            if ($primaryKey && $primaryKeyCol === null) {
                return redirect()->route('file.index')->with('error', 'Không tìm thấy cột khóa chính trong file Excel.');
            }

            $generatedDocs = session('generated_doc_files', []);
            $docCount = count($generatedDocs[$docIndex] ?? []) + 1;

            for ($row = 2; $row <= $highestRow; $row++) {
                $rowData = [];
                $primaryKeyValue = null;
                for ($col = 1; $col <= $highestColumnIndex; $col++) {
                    $value = $worksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
                    $rowData[$headerRow[$col]] = $value ? trim($value) : '';
                    if ($col === $primaryKeyCol) {
                        $primaryKeyValue = $rowData[$headerRow[$col]];
                    }
                }

                if ($primaryKey && empty($primaryKeyValue)) {
                    continue;
                }

                $templateProcessor = new TemplateProcessor($docPath);
                foreach ($relevantMappings as $mapping) {
                    $fieldName = $mapping['field']['field'];
                    $variable = $mapping['variable'];
                    $value = $rowData[$fieldName] ?? '';
                    $templateProcessor->setValue($variable, $value);
                }

                $outputFilename = basename($docPath, '.docx') . '_' . $docCount . '.docx';
                $outputPath = rtrim($outputFolder, '\\') . '\\' . $outputFilename;
                $templateProcessor->saveAs($outputPath);

                $generatedDocs[$docIndex][] = [
                    'filename' => $outputFilename,
                    'path' => $outputPath
                ];
                $docCount++;
            }

            session(['generated_doc_files' => $generatedDocs]);
            return redirect()->route('file.index')->with('success', 'Đã tạo các file tài liệu thành công.');

        } catch (\Exception $e) {
            Log::error('Lỗi khi tạo document: ' . $e->getMessage());
            return redirect()->route('file.index')->with('error', 'Không thể tạo tài liệu: ' . $e->getMessage());
        }
    }
}

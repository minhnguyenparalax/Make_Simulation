<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\Element\Text as TextElement;
use PhpOffice\PhpWord\Element\Table;
use Illuminate\Support\Facades\Log;

class VariablesDocController extends Controller
{
    public function getVariables($docIndex)
    {
        $docFiles = session('doc_files', []);

        if (!isset($docFiles[$docIndex])) {
            return redirect()->route('file.index')->with('error', 'File Doc không tồn tại trong danh sách.');
        }

        $filePath = $docFiles[$docIndex]['path'];
        if (!file_exists($filePath)) {
            return redirect()->route('file.index')->with('error', 'File không tồn tại tại: ' . $filePath);
        }

        try {
            $phpWord = IOFactory::load($filePath);
            $variables = [];

            foreach ($phpWord->getSections() as $section) {
                foreach ($section->getElements() as $element) {
                    if ($element instanceof Table) {
                        foreach ($element->getRows() as $row) {
                            foreach ($row->getCells() as $cell) {
                                foreach ($cell->getElements() as $cellElement) {
                                    if ($cellElement instanceof TextRun) {
                                        foreach ($cellElement->getElements() as $subElement) {
                                            if ($subElement instanceof TextElement) {
                                                $text = $subElement->getText();
                                                if ($text && is_string($text)) {
                                                    preg_match_all('/\{\{([^{}]+)\}\}/', $text, $matches);
                                                    if (!empty($matches[1])) {
                                                        $variables = array_merge($variables, $matches[1]);
                                                    }
                                                }
                                            }
                                        }
                                    } elseif ($cellElement instanceof TextElement) {
                                        $text = $cellElement->getText();
                                        if ($text && is_string($text)) {
                                            preg_match_all('/\{\{([^{}]+)\}\}/', $text, $matches);
                                            if (!empty($matches[1])) {
                                                $variables = array_merge($variables, $matches[1]);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    } elseif ($element instanceof TextRun) {
                        foreach ($element->getElements() as $subElement) {
                            if ($subElement instanceof TextElement) {
                                $text = $subElement->getText();
                                if ($text && is_string($text)) {
                                    preg_match_all('/\{\{([^{}]+)\}\}/', $text, $matches);
                                    if (!empty($matches[1])) {
                                        $variables = array_merge($variables, $matches[1]);
                                    }
                                }
                            }
                        }
                    } elseif ($element instanceof TextElement) {
                        $text = $element->getText();
                        if ($text && is_string($text)) {
                            preg_match_all('/\{\{([^{}]+)\}\}/', $text, $matches);
                            if (!empty($matches[1])) {
                                $variables = array_merge($variables, $matches[1]);
                            }
                        }
                    }
                }
            }

            $variables = array_unique(array_map('trim', $variables));

            $docVariables = session('doc_variables', []);
            $docVariables[$docIndex] = [
                'doc_name' => $docFiles[$docIndex]['name'],
                'variables' => $variables,
                'primary_key' => null,
            ];
            session(['doc_variables' => $docVariables]);

            return redirect()->back()->with('success', 'Đã lấy danh sách biến của file "' . $docFiles[$docIndex]['name'] . '" thành công.');
        } catch (\Exception $e) {
            Log::error('Lỗi khi đọc biến: ' . $e->getMessage());
            return redirect()->route('file.index')->with('error', 'Không thể đọc biến: ' . $e->getMessage());
        }
    }

    public function removeVariables($docIndex)
    {
        $docVariables = session('doc_variables', []);
        $mappings = session('mappings', []);

        if (isset($docVariables[$docIndex])) {
            $mappings = array_filter($mappings, fn($mapping) => $mapping['doc_index'] != $docIndex);
            unset($docVariables[$docIndex]);
            $generatedDocFiles = session('generated_doc_files', []);
            unset($generatedDocFiles[$docIndex]);

            session([
                'mappings' => array_values($mappings),
                'doc_variables' => $docVariables,
                'generated_doc_files' => $generatedDocFiles
            ]);

            return redirect()->back()->with('success', 'Đã xóa danh sách biến và khóa chính của file.');
        }

        return redirect()->back()->with('error', 'Danh sách biến không tồn tại.');
    }
}
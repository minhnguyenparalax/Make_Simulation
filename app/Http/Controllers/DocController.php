<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Style\Paragraph;
use Illuminate\Support\Facades\Log;

class DocController extends Controller
{
    public function addDoc(Request $request)
    {
        $request->validate([
            'file_path' => 'required|string',
        ]);

        $filePath = trim($request->input('file_path'), '"\'');
        $filePath = str_replace('/', '\\', $filePath);

        if (!file_exists($filePath)) {
            return back()->with('error', 'File không tồn tại tại: ' . $filePath);
        }

        $allowedExtensions = ['doc', 'docx'];
        $extension = strtolower(pathinfo($filePath, PATHINFO_EXTENSION));
        if (!in_array($extension, $allowedExtensions)) {
            return back()->with('error', 'Định dạng file không hợp lệ. Chỉ hỗ trợ .doc hoặc .docx.');
        }

        try {
            $phpWord = IOFactory::load($filePath, 'Word2007');
            $docFiles = session('doc_files', []);
            $docFiles[] = [
                'path' => $filePath,
                'name' => basename($filePath),
            ];

            session(['doc_files' => $docFiles]);

            return redirect()->route('file.index')->with('success', 'Đã thêm file Doc thành công: ' . $filePath);

        } catch (\Exception $e) {
            Log::error('Lỗi khi đọc file Doc: ' . $e->getMessage());
            return back()->with('error', 'Không thể đọc file Doc: Định dạng sai hoặc file hỏng.');
        }
    }

    public function removeDoc(Request $request)
    {
        $request->validate([
            'doc_index' => 'required|integer|min:0',
        ]);

        $docIndex = $request->input('doc_index');
        $docFiles = session('doc_files', []);
        $docVariables = session('doc_variables', []);
        $mappings = session('mappings', []);
        $generatedDocFiles = session('generated_doc_files', []);

        if (!isset($docFiles[$docIndex])) {
            return back()->with('error', 'File không tồn tại trong danh sách.');
        }

        $filePath = $docFiles[$docIndex]['path'];
        unset($docFiles[$docIndex]);
        $docFiles = array_values($docFiles);

        // Xóa các biến liên quan
        unset($docVariables[$docIndex]);
        // Xóa các mapping liên quan
        $mappings = array_filter($mappings, fn($mapping) => $mapping['doc_index'] != $docIndex);
        $mappings = array_values($mappings);
        // Xóa các file Doc đã tạo
        unset($generatedDocFiles[$docIndex]);

        session([
            'doc_files' => $docFiles,
            'doc_variables' => $docVariables,
            'mappings' => $mappings,
            'generated_doc_files' => $generatedDocFiles
        ]);

        return redirect()->route('file.index')->with('success', 'Đã xóa file Doc: ' . $filePath);
    }

    public function readDoc($docIndex)
    {
        $docFiles = session('doc_files', []);
        $excelFiles = session('excel_files', []);

        if (!isset($docFiles[$docIndex])) {
            return redirect()->route('file.index')->with('error', 'File không tồn tại trong danh sách.');
        }

        $filePath = $docFiles[$docIndex]['path'];
        if (!file_exists($filePath)) {
            return redirect()->route('file.index')->with('error', 'File không tồn tại tại: ' . $filePath);
        }

        try {
            $phpWord = IOFactory::load($filePath, 'Word2007');
            $content = '';

            foreach ($phpWord->getSections() as $section) {
                foreach ($section->getElements() as $element) {
                    if ($element instanceof TextRun) {
                        $paraStyle = $element->getParagraphStyle();
                        $alignment = $this->getAlignment($paraStyle);
                        $content .= "<p style=\"text-align: $alignment;\">";
                        foreach ($element->getElements() as $textElement) {
                            if (method_exists($textElement, 'getText')) {
                                $text = htmlspecialchars($textElement->getText());
                                $fontStyle = $textElement->getFontStyle();
                                $style = '';
                                if ($fontStyle) {
                                    if ($fontStyle->isBold()) {
                                        $style .= 'font-weight: bold;';
                                    }
                                    if ($fontStyle->isItalic()) {
                                        $style .= 'font-style: italic;';
                                    }
                                }
                                $content .= $style ? "<span style=\"$style\">$text</span>" : $text;
                            }
                        }
                        $content .= '</p>';
                    } elseif ($element instanceof Table) {
                        $content .= '<table class="table table-bordered">';
                        foreach ($element->getRows() as $row) {
                            $content .= '<tr>';
                            foreach ($row->getCells() as $cell) {
                                $content .= '<td>';
                                foreach ($cell->getElements() as $cellElement) {
                                    if ($cellElement instanceof TextRun) {
                                        $paraStyle = $cellElement->getParagraphStyle();
                                        $alignment = $this->getAlignment($paraStyle);
                                        $content .= "<div style=\"text-align: $alignment;\">";
                                        foreach ($cellElement->getElements() as $textElement) {
                                            if (method_exists($textElement, 'getText')) {
                                                $text = htmlspecialchars($textElement->getText());
                                                $fontStyle = $textElement->getFontStyle();
                                                $style = '';
                                                if ($fontStyle) {
                                                    if ($fontStyle->isBold()) {
                                                        $style .= 'font-weight: bold;';
                                                    }
                                                    if ($fontStyle->isItalic()) {
                                                        $style .= 'font-style: italic;';
                                                    }
                                                }
                                                $content .= $style ? "<span style=\"$style\">$text</span>" : $text;
                                            }
                                        }
                                        $content .= '</div>';
                                    }
                                }
                                $content .= '</td>';
                            }
                            $content .= '</tr>';
                        }
                        $content .= '</table>';
                    } elseif (method_exists($element, 'getParagraphStyle')) {
                        $paraStyle = $element->getParagraphStyle();
                        $alignment = $this->getAlignment($paraStyle);
                        $content .= "<p style=\"text-align: $alignment;\">";
                        if (method_exists($element, 'getText')) {
                            $text = htmlspecialchars($element->getText());
                            $content .= nl2br($text);
                        }
                        $content .= '</p>';
                    }
                    $content .= '<br>';
                }
            }

            if (empty($content)) {
                return redirect()->route('file.index')->with('error', 'Không tìm thấy nội dung trong file Doc.');
            }

            return view('file_reader', [
                'excelFiles' => $excelFiles,
                'docFiles' => $docFiles,
                'docContent' => $content,
                'currentDocIndex' => $docIndex,
                'success' => 'Đã đọc file Doc "' . $docFiles[$docIndex]['name'] . '" thành công.'
            ]);

        } catch (\Exception $e) {
            Log::error('Lỗi khi đọc file Doc: ' . $e->getMessage());
            return redirect()->route('file.index')->with('error', 'Không thể đọc file Doc: Định dạng sai hoặc file hỏng.');
        }
    }

    private function getAlignment($paraStyle)
    {
        if ($paraStyle instanceof \PhpOffice\PhpWord\Style\Paragraph) {
            $alignment = $paraStyle->getAlignment();
            switch ($alignment) {
                case 'center':
                    return 'center';
                case 'right':
                    return 'right';
                case 'justify':
                case 'both':
                    return 'justify';
                default:
                    return 'left';
            }
        }
        return 'left';
    }
}
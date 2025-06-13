<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\File;

class FolderController extends Controller
{
    public function showForm()
    {
        return view('folder_form');
    }

    public function createFolder(Request $request)
    {
        $request->validate([
            'path' => 'required|string',
            'folder_name' => 'required|string'
        ]);

        $basePath = trim($request->input('path'), "\"' "); // Xóa dấu ngoặc hoặc dấu cách
        $folderName = trim($request->input('folder_name'), "\"' ");

        $fullPath = rtrim($basePath, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $folderName;

        if (File::exists($fullPath)) {
            return back()->with('error', 'Thư mục đã tồn tại: ' . $fullPath);
        }

        try {
            File::makeDirectory($fullPath, 0755, true);
            return back()->with('success', 'Đã tạo thư mục thành công: ' . $fullPath);
        } catch (\Exception $e) {
            return back()->with('error', 'Lỗi khi tạo thư mục: ' . $e->getMessage());
        }
    }
}
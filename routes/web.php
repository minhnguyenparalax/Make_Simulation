<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExcelController;
use App\Http\Controllers\DocController;
use App\Http\Controllers\FieldExcelController;
use App\Http\Controllers\VariablesDocController;
use App\Http\Controllers\FieldMappingController;
use App\Http\Controllers\GenerateDocController;
use App\Http\Controllers\FileController;

Route::get('/', function () {
    return view('file_reader', [
        'excelFiles' => session('excel_files', []),
        'docFiles' => session('doc_files', []),
    ]);
})->name('file.index');

Route::post('/excel/add', [ExcelController::class, 'addExcel'])->name('excel.addExcel');
Route::delete('/excel/remove/{fileIndex}', [ExcelController::class, 'removeExcel'])->name('excel.removeExcel');
Route::get('/excel/read/{fileIndex}/{sheetIndex}', [ExcelController::class, 'readSheet'])->name('excel.readSheet');

Route::post('/doc/add', [DocController::class, 'addDoc'])->name('doc.addDoc');
Route::post('/doc/remove', [DocController::class, 'removeDoc'])->name('doc.removeDoc');
Route::get('/doc/read/{docIndex}', [DocController::class, 'readDoc'])->name('doc.readDoc');

Route::get('/excel/fields/{fileIndex}/{sheetIndex}', [FieldExcelController::class, 'getFields'])->name('excel.fields');
Route::get('/excel/fields/remove/{fileIndex}/{sheetIndex}', [FieldExcelController::class, 'removeFields'])->name('excel.removeFields');

Route::get('/doc/variables/{docIndex}', [VariablesDocController::class, 'getVariables'])->name('doc.variables');
Route::get('/doc/variables/remove/{docIndex}', [VariablesDocController::class, 'removeVariables'])->name('doc.removeVariables');
Route::post('/doc/map-variable', [FieldMappingController::class, 'mapVariable'])->name('doc.mapVariable');
Route::post('/doc/remove-mapping', [FieldMappingController::class, 'removeMapping'])->name('doc.removeMapping');

Route::post('/doc/set-primary-key', [GenerateDocController::class, 'setPrimaryKey'])->name('doc.setPrimaryKey');
Route::post('/doc/set-output-folder', [GenerateDocController::class, 'setOutputFolder'])->name('doc.setOutputFolder');
Route::get('/doc/generate/{docIndex}', [GenerateDocController::class, 'generateDoc'])->name('doc.generate');

Route::get('/clear-session', [FileController::class, 'clearSession'])->name('clear.session');
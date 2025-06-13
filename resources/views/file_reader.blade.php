<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel & Doc Viewer</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        body { font-size: 0.9rem; }
        .container { max-width: 1000px; padding: 0 10px; }
        .status-column { font-weight: bold; }
        .table-smaller { font-size: 0.85rem; }
        .table-smaller th, .table-smaller td { padding: 0.4rem; line-height: 1.1; }
        .table-smaller th { background-color: #f1f1f1; }
        .sheet-list, .doc-list { list-style: none; margin: 0; padding: 0; }
        .doc-content { border: 1px solid #dee2e6; padding: 0.8rem; margin-top: 0.8rem; background-color: #f8f9fa; white-space: pre-wrap; }
        .doc-content p { margin: 0 0 0.4rem 0; }
        .doc-content table { width: 100%; margin-bottom: 0.8rem; }
        .doc-content table th, .doc-content table td { padding: 0.4rem; vertical-align: top; }
        .section-title { margin-bottom: 0.8rem; font-size: 1.2rem; }
        .form-section, .list-section { margin-bottom: 1.5rem; }
        .sheet-item, .doc-item { display: flex; align-items: center; gap: 0.5rem; padding: 0.2rem 0; }
        .excel-item { display: flex; justify-content: space-between; align-items: center; gap: 1rem; }
        .eye-btn, .map-btn, .toggle-btn, .map-v-btn, .remove-map-btn {
            font-size: 1.1rem; padding: 0.3rem; border: none; border-radius: 50%;
            cursor: pointer; transition: transform 0.2s, box-shadow 0.2s, background-color 0.2s;
            min-width: 1.8rem; text-align: center;
        }
        .eye-btn { color: #0d6efd; background-color: #e9ecef; }
        .map-btn { color: #198754; background-color: #e9ecef; }
        .toggle-btn { color: #6c757d; background-color: #e9ecef; }
        .map-v-btn { color: #6f42c1; background-color: #e9ecef; }
        .remove-map-btn { color: #dc3545; background-color: #e9ecef; }
        .eye-btn:hover { transform: scale(1.1); box-shadow: 0 2px 4px rgba(0,0,0,0.2); background-color: #d0e7ff; color: #0056b3; }
        .map-btn:hover { transform: scale(1.1); box-shadow: 0 2px 4px rgba(0,0,0,0.2); background-color: #d4edda; color: #146c43; }
        .toggle-btn:hover { transform: scale(1.1); box-shadow: 0 2px 4px rgba(0,0,0,0.2); background-color: #e2e6ea; color: #495057; }
        .map-v-btn:hover { transform: scale(1.1); box-shadow: 0 2px 4px rgba(0,0,0,0.2); background-color: #e9d8fd; color: #5a32a3; }
        .remove-map-btn:hover { transform: scale(1.1); box-shadow: 0 2px 4px rgba(0,0,0,0.2); background-color: #f8d7da; color: #b02a37; }
        .sheet-actions, .doc-actions { display: flex; align-items: center; gap: 0.4rem; }
        .doc-item .delete-form { margin-left: auto; }
        .sheet-name, .doc-name { flex: 1; }
        .fields-table { margin-top: 0.8rem; font-size: 0.85rem; width: 100%; }
        .fields-scroll { overflow-x: auto; overflow-y: auto; max-height: 150px; margin-bottom: 0.8rem; }
        .variables-list { margin-top: 0.8rem; padding-left: 0.8rem; }
        .variables-scroll { overflow-y: auto; max-height: 400px; margin-bottom: 0.8rem; }
        .field-section, .variable-section {
            margin-top: 1rem; padding: 0.8rem; border: 1px solid #dee2e6;
            border-radius: 0.25rem; background-color: #fff;
        }
        .fields-scroll::-webkit-scrollbar, .variables-scroll::-webkit-scrollbar {
            width: 6px; height: 6px;
        }
        .fields-scroll::-webkit-scrollbar-thumb, .variables-scroll::-webkit-scrollbar-thumb {
            background-color: #adb5bd; border-radius: 4px;
        }
        .fields-scroll::-webkit-scrollbar-track, .variables-scroll::-webkit-scrollbar-track {
            background-color: #f1f1f1;
        }
        .sheet-list.hidden, .doc-list.hidden, .mapping-list.hidden { display: none; }
        .excel-header { display: flex; align-items: center; gap: 0.4rem; }
        .variable-item { display: flex; align-items: center; gap: 0.4rem; }
        .dropdown-menu {
            max-height: 300px; overflow-y: auto; width: 300px; font-size: 0.85rem;
        }
        .dropdown-item { padding: 0.4rem 1rem; }
        .dropdown-header { font-weight: bold; color: #343a40; background-color: #e9ecef; padding: 0.5rem 1rem; }
        .mapping-toggle-btn { margin-bottom: 0.8rem; font-size: 0.85rem; }
        .btn-sm { font-size: 0.7rem; padding: 0.2rem 0.5rem; }
        .mapping-field, .mapping-variable { color: #0d6efd; font-weight: 500; }
        .mapping-file-info { color: #6c757d; font-size: 0.85rem; }
        .mapping-item { margin-bottom: 0.5rem; }
    </style>
</head>
<body>
    <div class="container mt-4">
        <!-- Debug Data (có thể xóa sau khi debug xong) -->
        @if (false)
            @if (!empty($excelFiles))
                <pre>Excel Files: {{ print_r($excelFiles, true) }}</pre>
            @else
                <p class="text-warning">Không có file Excel nào.</p>
            @endif
            @if (session('sheet_fields'))
                <pre>Sheet Fields: {{ print_r(session('sheet_fields'), true) }}</pre>
            @else
                <p class="text-warning">Không có dữ liệu sheet_fields trong session.</p>
            @endif
            @if (session('mappings'))
                <pre>Mappings: {{ print_r(session('mappings'), true) }}</pre>
            @else
                <p class="text-warning">Không có mappings trong session.</p>
            @endif
            @if (!empty($docFiles))
                <pre>Doc Files: {{ print_r($docFiles, true) }}</pre>
            @else
                <p class="text-warning">Không có file Doc nào.</p>
            @endif
        @endif

        <div class="row">
            <!-- Cột trái: Đọc File Excel và Danh sách Excel -->
            <div class="col-12 col-md-6">
                <div class="form-section">
                    <h2 class="section-title">Đọc File Excel</h2>
                    <form action="{{ route('excel.addExcel') }}" method="POST">
                        @csrf
                        <div class="mb-2">
                            <label for="excel_file_path" class="form-label">Nhập Đường Dẫn File Excel (.xlsx hoặc .xls)</label>
                            <input type="text" class="form-control form-control-sm" id="excel_file_path" name="file_path"
                                   placeholder="VD: F:\00 Template\2025 Project Management.xlsx"
                                   value="{{ old('file_path') }}" required>
                            <small class="form-text text-muted">Đường dẫn có thể có hoặc không có dấu ngoặc kép</small>
                        </div>
                        <button type="submit" class="btn btn-primary btn-sm">Thêm Excel</button>
                    </form>
                </div>

                @if (!empty($excelFiles))
                    <div class="list-section">
                        <h3 class="section-title">Danh Sách Excel</h3>
                        <ul class="list-group">
                            @foreach ($excelFiles as $fileIndex => $file)
                                <li class="list-group-item">
                                    <div class="excel-item">
                                        <div>
                                            <div class="excel-header">
                                                <button type="button" class="toggle-btn" data-target="sheet-list-{{ $fileIndex }}">
                                                    <i class="bi bi-chevron-down"></i>
                                                </button>
                                                <strong>{{ $file['name'] }}</strong>
                                            </div>
                                            <ul class="sheet-list hidden" id="sheet-list-{{ $fileIndex }}">
                                                @foreach ($file['sheets'] as $sheetIndex => $sheetName)
                                                    <li class="sheet-item">
                                                        <div class="sheet-actions">
                                                            <form action="{{ route('excel.readSheet', [$fileIndex, $sheetIndex]) }}" method="GET" class="d-inline">
                                                                <button type="submit" class="eye-btn" title="Kiểm tra"><i class="bi bi-eye"></i></button>
                                                            </form>
                                                            <form action="{{ route('excel.fields', [$fileIndex, $sheetIndex]) }}" method="GET" class="d-inline">
                                                                <button type="submit" class="map-btn" title="Chuẩn bị mapping"><i class="bi bi-diagram-3"></i></button>
                                                            </form>
                                                        </div>
                                                        <span class="sheet-name {{ isset($currentFileIndex) && isset($currentSheetIndex) && $currentFileIndex == $fileIndex && $currentSheetIndex == $sheetIndex ? 'fw-bold' : '' }}">
                                                            {{ $sheetName }}
                                                        </span>
                                                    </li>
                                                @endforeach
                                            </ul>
                                        </div>
                                        <form action="{{ route('excel.removeExcel', $fileIndex) }}" method="POST">
                                            @csrf
                                            @method('DELETE')
                                            <button type="submit" class="btn btn-danger btn-sm">Xóa Excel</button>
                                        </form>
                                    </div>
                                </li>
                            @endforeach
                        </ul>

                        @if (session('sheet_fields'))
                            @foreach (session('sheet_fields') as $fIndex => $sheets)
                                @foreach ($sheets as $sIndex => $sheetData)
                                    <div class="field-section">
                                        <div class="d-flex justify-content-between align-items-center mb-2">
                                            <h4 class="section-title">Trường {{ $sheetData['sheet_name'] }}</h4>
                                            <form action="{{ route('excel.removeFields', [$fIndex, $sIndex]) }}" method="GET">
                                                <button type="submit" class="btn btn-danger btn-sm">Xóa danh sách này</button>
                                            </form>
                                        </div>
                                        <div class="fields-scroll">
                                            <table class="table table-bordered fields-table">
                                                <thead>
                                                    <tr>
                                                        @foreach ($sheetData['fields'] as $field)
                                                            <th>{{ $field }}</th>
                                                        @endforeach
                                                        @if (empty($sheetData['fields']))
                                                            <th>Chưa có trường nào</th>
                                                        @endif
                                                    </tr>
                                                </thead>
                                            </table>
                                        </div>
                                    </div>
                                @endforeach
                            @endforeach
                        @endif
                    </div>
                @endif
            </div>

            <!-- Cột phải: Đọc File Doc và Danh sách Doc -->
            <div class="col-12 col-md-6">
                <div class="form-section">
                    <h2 class="section-title">Đọc File Doc</h2>
                    <form action="{{ route('doc.addDoc') }}" method="POST">
                        @csrf
                        <div class="mb-2">
                            <label for="doc_file_path" class="form-label">Nhập Đường Dẫn File Doc (.doc hoặc .docx)</label>
                            <input type="text" class="form-control form-control-sm" id="doc_file_path" name="file_path"
                                   placeholder="VD: F:\Documents\Report.docx"
                                   value="{{ old('file_path') }}" required>
                            <small class="form-text text-muted">Đường dẫn có thể có hoặc không có dấu ngoặc kép</small>
                        </div>
                        <button type="submit" class="btn btn-primary btn-sm">Thêm Doc</button>
                    </form>
                </div>

                @if (!empty($docFiles))
                    <div class="list-section">
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <h3 class="section-title">Danh Sách Doc</h3>
                            <button type="button" class="toggle-btn" data-target="doc-list">
                                <i class="bi bi-chevron-down"></i>
                            </button>
                        </div>
                        <ul class="list-group doc-list hidden" id="doc-list">
                            @foreach ($docFiles as $docIndex => $doc)
                                <li class="list-group-item">
                                    <div class="doc-item">
                                        <div class="doc-actions">
                                            <form action="{{ route('doc.readDoc', [$docIndex]) }}" method="GET" class="d-inline">
                                                <button type="submit" class="eye-btn"><i class="bi bi-eye"></i></button>
                                            </form>
                                            <form action="{{ route('doc.variables', [$docIndex]) }}" method="GET" class="d-inline">
                                                <button type="submit" class="map-btn"><i class="bi bi-diagram-3"></i></button>
                                            </form>
                                            <span class="doc-name {{ isset($currentDocIndex) && $currentDocIndex === $docIndex ? 'fw-bold' : '' }}">
                                                {{ $doc['name'] }}
                                            </span>
                                        </div>
                                        <form action="{{ route('doc.removeDoc') }}" method="POST" class="delete-form">
                                            @csrf
                                            <input type="hidden" name="doc_index" value="{{ $docIndex }}">
                                            <button type="submit" class="btn btn-danger btn-sm">Xóa Doc</button>
                                        </form>
                                    </div>
                                </li>
                            @endforeach
                        </ul>

                        <button type="button" class="btn btn-primary btn-sm mapping-toggle-btn" data-target="mapping-list">Hiện Danh Sách Mapping</button>

                        <div class="mapping-list hidden" id="mapping-list">
                            @if (session('mappings') && !empty(session('mappings')))
                                <h4 class="section-title">Danh sách Mapping Trường-Biến</h4>
                                @foreach (session('mappings') as $index => $mapping)
                                    @php
                                        $isValidMapping = isset($mapping['field']['file_index']) &&
                                                          isset($mapping['field']['sheet_index']) &&
                                                          isset($mapping['field']['field']) &&
                                                          isset($mapping['doc_index']) &&
                                                          isset($mapping['variable']) &&
                                                          isset($excelFiles[$mapping['field']['file_index']]['name']) &&
                                                          isset($docFiles[$mapping['doc_index']]['name']);
                                    @endphp
                                    @if ($isValidMapping)
                                        <div class="mapping-item">
                                            <strong>{{ $index + 1 }}. </strong>
                                            <div class="mapping-file-info">
                                                ({{ $excelFiles[$mapping['field']['file_index']]['name'] }}/Sheet {{ $mapping['field']['sheet_index'] + 1 }}) ->
                                                ({{ $docFiles[$mapping['doc_index']]['name'] }})
                                            </div>
                                            <div>
                                                <span class="mapping-field">{{ $mapping['field']['field'] }}</span> ->
                                                <span class="mapping-variable">{{ $mapping['variable'] }}</span>
                                            </div>
                                        </div>
                                    @else
                                        <p class="text-danger">Lỗi: Dữ liệu mapping không hợp lệ tại chỉ số {{ $index }}</p>
                                    @endif
                                @endforeach
                            @else
                                <p>Chưa có mapping nào được thiết lập.</p>
                            @endif
                        </div>

                        @if (session('doc_variables'))
                            @foreach (session('doc_variables') as $dIndex => $docData)
                                <div class="variable-section">
                                    <div class="d-flex justify-content-between align-items-center mb-2">
                                        <h4 class="section-title">Biến {{ $docData['doc_name'] }}</h4>
                                        <form action="{{ route('doc.removeVariables', [$dIndex]) }}" method="GET">
                                            <button type="submit" class="btn btn-danger btn-sm">Xóa danh sách này</button>
                                        </form>
                                    </div>
                                    <div class="variables-scroll">
                                        <ul class="variables-list">
                                            @foreach ($docData['variables'] as $variable)
                                                <li class="variable-item">
                                                    <span>{{ $variable }}
                                                        @php
                                                            $mappings = session('mappings', []);
                                                            $mappedField = collect($mappings)->firstWhere(fn($m) => $m['doc_index'] == $dIndex && $m['variable'] == $variable);
                                                        @endphp
                                                        @if ($mappedField)
                                                            -> {{ $mappedField['field']['field'] }}
                                                            <form action="{{ route('doc.removeMapping') }}" method="POST" class="d-inline">
                                                                @csrf
                                                                <input type="hidden" name="doc_index" value="{{ $dIndex }}">
                                                                <input type="hidden" name="variable" value="{{ $variable }}">
                                                                <button type="submit" class="remove-map-btn">
                                                                    <i class="bi bi-x-circle"></i>
                                                                </button>
                                                            </form>
                                                        @endif
                                                    </span>
                                                    <div class="dropdown">
                                                        <button type="button" class="map-v-btn dropdown-toggle" data-bs-toggle="dropdown" data-bs-auto-close="outside">
                                                            <i class="bi bi-link-45deg"></i>
                                                        </button>
                                                        <ul class="dropdown-menu">
                                                            @if (session('sheet_fields'))
                                                                @foreach (session('sheet_fields') as $fIndex => $sheets)
                                                                    @foreach ($sheets as $sIndex => $sFields)
                                                                        <li class="dropdown-header">
                                                                            {{ $excelFiles[$fIndex]['name'] ?? 'Unknown File' }}/{{ $sFields['sheet_name'] }}
                                                                        </li>
                                                                        @foreach ($sFields['fields'] as $field)
                                                                            @php
                                                                                $isFieldUsed = collect($mappings)->contains(fn($m) => 
                                                                                    $m['doc_index'] == $dIndex && 
                                                                                    $m['field']['file_index'] == $fIndex && 
                                                                                    $m['field']['sheet_index'] == $sIndex && 
                                                                                    $m['field']['field'] == $field
                                                                                );
                                                                            @endphp
                                                                            <li>
                                                                                <form action="{{ route('doc.mapVariable') }}" method="POST">
                                                                                    @csrf
                                                                                    <input type="hidden" name="doc_index" value="{{ $dIndex }}">
                                                                                    <input type="hidden" name="variable" value="{{ $variable }}">
                                                                                    <input type="hidden" name="file_index" value="{{ $fIndex }}">
                                                                                    <input type="hidden" name="sheet_index" value="{{ $sIndex }}">
                                                                                    <input type="hidden" name="field" value="{{ $field }}">
                                                                                    <button type="submit" class="dropdown-item" {{ $isFieldUsed ? 'disabled' : '' }}>
                                                                                        {{ $field }} {{ $isFieldUsed ? '(Đã sử dụng trong Doc này)' : '' }}
                                                                                    </button>
                                                                                </form>
                                                                            </li>
                                                                        @endforeach
                                                                    @endforeach
                                                                @endforeach
                                                            @else
                                                                <li class="dropdown-item disabled">Chưa có trường nào để mapping</li>
                                                            @endif
                                                        </ul>
                                                    </div>
                                                </li>
                                            @endforeach
                                        </ul>
                                    </div>
                                </div>
                            @endforeach
                        @endif
                    </div>
                @endif
            </div>
        </div>

        @if (session()->has('error'))
            <div class="row mt-3">
                <div class="col-12">
                    <div class="alert alert-danger">
                        {{ session('error') }}
                    </div>
                </div>
            </div>
        @endif

        @if (session()->has('success'))
            <div class="row mt-3">
                <div class="col-12">
                    <div class="alert alert-success">
                        {{ session('success') }}
                    </div>
                </div>
            </div>
        @endif

        @if (isset($data) && !empty($data))
            <div class="row mt-4">
                <div class="col-12">
                    <h3 class="section-title">Dữ Liệu: {{ $excelFiles[$currentFileIndex]['name'] }} / {{ $excelFiles[$currentFileIndex]['sheets'][$currentSheetIndex] }}</h3>
                    <table class="table table-bordered mt-2 table-smaller">
                        <thead>
                            <tr>
                                @foreach ($data[0] as $index => $cell)
                                    @if ($cell['colspan'] !== 0)
                                        <th class="{{ isset($statusColumnIndex) && $index == $statusColumnIndex ? 'status-column' : '' }}"
                                            @if ($cell['colspan'] > 1) colspan="{{ $cell['colspan'] }}" @endif>
                                            {{ $cell['value'] !== '' ? $cell['value'] : '' }}
                                        </th>
                                    @endif
                                @endforeach
                            </tr>
                        </thead>
                        <tbody>
                            @foreach (array_slice($data, 1) as $row)
                                <tr>
                                    @foreach ($row as $colIndex => $cell)
                                        @if ($cell['colspan'] !== 0)
                                            <td class="{{ isset($statusColumnIndex) && $colIndex == $statusColumnIndex ? 'status-column' : '' }}"
                                                @if ($cell['colspan'] > 1) colspan="{{ $cell['colspan'] }}" @endif>
                                                {{ $cell['value'] !== '' ? $cell['value'] : '' }}
                                            </td>
                                        @endif
                                    @endforeach
                                </tr>
                            @endforeach
                        </tbody>
                    </table>
                </div>
            </div>
        @endif

        @if (isset($docContent))
            <div class="row mt-4">
                <div class="col-12">
                    <h3 class="section-title">Nội Dung: {{ $docFiles[$currentDocIndex]['name'] }}</h3>
                    <div class="doc-content">
                        {!! $docContent !!}
                    </div>
                </div>
            </div>
        @endif
    </div>

    <script>
        document.querySelectorAll('.toggle-btn').forEach(button => {
            const targetId = button.getAttribute('data-target');
            const target = document.getElementById(targetId);
            const icon = button.querySelector('i');
            const isExpanded = localStorage.getItem(`toggle-${targetId}`) === 'true';
            if (isExpanded) {
                target.classList.remove('hidden');
                icon.classList.remove('bi-chevron-down');
                icon.classList.add('bi-chevron-up');
            } else {
                target.classList.add('hidden');
                icon.classList.remove('bi-chevron-up');
                icon.classList.add('bi-chevron-down');
            }
            button.addEventListener('click', () => {
                const isHidden = target.classList.contains('hidden');
                if (isHidden) {
                    target.classList.remove('hidden');
                    icon.classList.remove('bi-chevron-down');
                    icon.classList.add('bi-chevron-up');
                    localStorage.setItem(`toggle-${targetId}`, 'true');
                } else {
                    target.classList.add('hidden');
                    icon.classList.remove('bi-chevron-up');
                    icon.classList.add('bi-chevron-down');
                    localStorage.setItem(`toggle-${targetId}`, 'false');
                }
            });
        });
        document.querySelectorAll('.mapping-toggle-btn').forEach(button => {
            button.addEventListener('click', () => {
                const target = document.getElementById('mapping-list');
                const isHidden = target.classList.contains('hidden');
                if (isHidden) {
                    target.classList.remove('hidden');
                    button.textContent = 'Ẩn Danh Sách Mapping';
                } else {
                    target.classList.add('hidden');
                    button.textContent = 'Hiện Danh Sách Mapping';
                }
            });
        });
    </script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
</body>
</html>
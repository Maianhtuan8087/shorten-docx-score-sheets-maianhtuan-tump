# shorten-docx-score-sheets-maianhtuan-tump

Web tĩnh chuyển file đáp án DOCX thành phiếu chấm rút gọn ngay trong trình duyệt.

## Tính năng

- Upload hoặc kéo thả file DOCX.
- Rút gọn nội dung theo quy tắc 5 từ đầu + 3 từ cuối.
- Tạo phiếu chấm 4 cột bằng template đã duyệt.
- Hiển thị báo cáo QC và xem trước nội dung ở bên phải.
- Tải file DOCX kết quả trực tiếp.
- File được xử lý trong bộ nhớ trình duyệt, không gửi tới backend.

## Chạy tại máy

```powershell
python -m http.server 4173
```

Mở `http://localhost:4173`. Lần đầu cần có Internet để tải Pyodide và `python-docx`.

## Kiểm thử

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest
```

## Triển khai

Workflow GitHub Pages nằm tại `.github/workflows/pages.yml`. Bật Pages với nguồn **GitHub Actions** trong phần Settings của repository.

## Tác giả

Mai Anh Tuấn TUMP  
Social Medicine Public Health

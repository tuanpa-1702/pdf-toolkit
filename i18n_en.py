# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
"""Bản dịch tiếng Anh (khoá = chữ tiếng Việt trong mã nguồn). Sinh bằng công cụ, sửa tay được.
PATTERNS: mẫu cho chữ ghép động, {0} {1}... là phần thay đổi."""

EXACT = {'   Mở PDF': '   Open PDF',
 '  (file đang có thay đổi chưa lưu - kết quả tính trên bản đang sửa)': '  (file has unsaved changes - results are '
                                                                        'based on the current edits)',
 '  Chạy': '  Run',
 '  Dừng': '  Stop',
 '  Tất cả công cụ': '  All tools',
 '  — ĐÃ HẾT HẠN, không ký được': '  — EXPIRED, cannot sign',
 ' (Không tìm thấy - chữ có thể là ảnh scan)': ' (Not found - text may be a scanned image)',
 ' - không phát hiện lỗi': ' - no errors detected',
 '(Không tìm thấy chứng thư số)': '(Digital certificate not found)',
 '(chưa chọn)': '(not selected)',
 '(file không có lớp)': '(file has no layers)',
 '(file không có đính kèm)': '(file has no attachments)',
 '(không có)': '(none)',
 '(không tìm thấy Office)': '(Office not found)',
 ');;Tất cả (*.*)': ');;All (*.*)',
 '0 = tự co chữ vừa ô': '0 = auto-fit text to the field',
 '0% = đen trắng': '0% = black and white',
 '1 | Chương 1 | 1\n2 | Mục 1.1 | 3\n1 | Chương 2 | 10': '1 | Chapter 1 | 1\n2 | Section 1.1 | 3\n1 | Chapter 2 | 10',
 '150 dpi - nhẹ, đủ đọc': '150 dpi - small, readable',
 '200 dpi - khuyên dùng': '200 dpi - recommended',
 '300 dpi - nét, để OCR / in': '300 dpi - sharp, for OCR / printing',
 '300 dpi là tối ưu. Chữ rất nhỏ (< 8pt) dùng 400': '300 dpi is optimal. For very small text (< 8pt) use 400',
 '600 dpi - rất nét (chậm, file nặng)': '600 dpi - very sharp (slow, large file)',
 '90° phải': '90° right',
 '90° trái': '90° left',
 '96 = xem màn hình, 150 = thường dùng, 300 = in ấn': '96 = screen viewing, 150 = typical use, 300 = printing',
 '< 1 làm tối vùng giữa, > 1 làm sáng': '< 1 darkens midtones, > 1 brightens',
 "<br><i style='color:#64748b'>Bấm để xem / sửa ghi chú</i>": "<br><i style='color:#64748b'>Click to view / edit "
                                                              'note</i>',
 '<p>(Không có nội dung)</p>': '<p>(No content)</p>',
 'Add-in Office: tab PDF Toolkit trong Word, Excel, PowerPoint': 'Office add-in: PDF Toolkit tab in Word, Excel, '
                                                                 'PowerPoint',
 'Add-in Office: tab PDF Toolkit trong Word, Excel, PowerPoint, Outlook': 'Office add-in: PDF Toolkit tab in Word, '
                                                                          'Excel, PowerPoint, Outlook',
 'Biến PDF / ảnh scan thành PDF tìm kiếm, bôi đen, copy được chữ (tiếng Việt + tiếng Anh), giữ nguyên hình scan.': 'Turn '
                                                                                                                   'a '
                                                                                                                   'PDF '
                                                                                                                   '/ '
                                                                                                                   'scanned '
                                                                                                                   'image '
                                                                                                                   'into '
                                                                                                                   'a '
                                                                                                                   'searchable '
                                                                                                                   'PDF '
                                                                                                                   'with '
                                                                                                                   'text '
                                                                                                                   'you '
                                                                                                                   'can '
                                                                                                                   'highlight '
                                                                                                                   'and '
                                                                                                                   'copy '
                                                                                                                   '(Vietnamese '
                                                                                                                   '+ '
                                                                                                                   'English), '
                                                                                                                   'while '
                                                                                                                   'keeping '
                                                                                                                   'the '
                                                                                                                   'original '
                                                                                                                   'scanned '
                                                                                                                   'image.',
 'Biến chữ thành đường vẽ (vector): hiển thị / in đúng mọi máy, không còn cần font, không copy chữ được.': 'Turn '
                                                                                                           'text '
                                                                                                           'into '
                                                                                                           'vector '
                                                                                                           'outlines: '
                                                                                                           'displays '
                                                                                                           '/ prints '
                                                                                                           'the same '
                                                                                                           'on every '
                                                                                                           'machine, '
                                                                                                           'no fonts '
                                                                                                           'needed, '
                                                                                                           'text can '
                                                                                                           'no '
                                                                                                           'longer '
                                                                                                           'be '
                                                                                                           'copied.',
 'Biến các địa chỉ web / email dạng chữ trong file thành liên kết bấm được': 'Turn web / email addresses written as '
                                                                             'text into clickable links',
 'Biến mọi trang thành ảnh: không chọn / sao chép / sửa chữ được nữa (giống bản scan sạch).': 'Turn every page into '
                                                                                              'an image: text can no '
                                                                                              'longer be selected / '
                                                                                              'copied / edited (like '
                                                                                              'a clean scan).',
 'Biến ô form, ghi chú, chữ ký thành nội dung cố định của trang (không sửa được nữa).': 'Flatten form fields, notes, '
                                                                                        'and signatures into fixed '
                                                                                        'page content (no longer '
                                                                                        'editable).',
 'Bookmark đầu tiên phải là cấp 1': 'The first bookmark must be level 1',
 'BÚT DẠ QUANG: kéo chuột để tô': 'FREE HIGHLIGHT: drag to highlight',
 'BÚT VẼ: kéo chuột để vẽ tự do (Màu & nét bút ở tab Khác)': 'PENCIL: drag to draw freehand (Pen color & width in '
                                                             'the Other tab)',
 'BÚT VẼ: kéo chuột để vẽ tự do (Màu & nét bút ở tab Nhận xét)': 'PENCIL: drag to draw freehand (Pen color & width '
                                                                 'in the Comment tab)',
 'Bút dạ quang': 'Free highlight',
 'Bút vẽ': 'Pencil',
 'Bạn chỉ bấm một điểm nên cần nhập chữ hiển thị (hoặc kéo vùng trên chữ có sẵn).': 'You clicked only a single '
                                                                                    'point, so you need to enter '
                                                                                    'display text (or drag an area '
                                                                                    'over existing text).',
 'BẢN CHÍNH': 'ORIGINAL',
 'BẢN NHÁP': 'DRAFT',
 'BẢN SAO': 'COPY',
 'Bản PDF Toolkit đã cài đang mở - hãy đóng tất cả cửa sổ PDF Toolkit rồi thử lại.': 'The installed PDF Toolkit is '
                                                                                     'running - close all PDF '
                                                                                     'Toolkit windows and try again.',
 'Bản exe này không kèm add-in Office.': 'This exe does not include the Office add-in.',
 'Bản quyền': 'Copyright',
 'Bản sao không còn hợp lệ.': 'The copy is no longer valid.',
 'Bảng trang': 'Page panel',
 'Bảo mật && Ký số': 'Security && Digital Signature',
 'Bảo mật, Form & Tối ưu': 'Security, Forms & Optimization',
 'Bấm chọn khung chữ / ảnh / con dấu / khối vẽ trên trang trước': 'Click a text box / image / stamp / shape on the '
                                                                  'page first',
 'Bấm chọn ảnh / chữ ký / con dấu / khối vẽ trên trang trước': 'Click an image / signature / stamp / shape on the '
                                                               'page first',
 'Bấm từng điểm, bấm đúp / Enter / chuột phải để kết thúc': 'Click each point; double-click / Enter / right-click to '
                                                            'finish',
 'Bấm từng đỉnh, bấm đúp / Enter / chuột phải để kết thúc': 'Click each vertex; double-click / Enter / right-click '
                                                            'to finish',
 'Bấm vào chỗ cần gõ chữ (không viền, không nền) - vd điền vào form giấy': 'Click where to type (no border, no fill) '
                                                                           '- e.g. fill in a paper form',
 'Bấm vào ghi chú / nét vẽ / hình để xoá': 'Click a comment / drawing / shape to delete it',
 'Bấm vào một dòng chữ để sửa (đổi font, đậm, nghiêng, màu...)': 'Click a line of text to edit (change font, bold, '
                                                                 'italic, color...)',
 'Bấm vào trang để đặt con dấu (hoặc kéo khung để chọn kích thước). Sau khi đặt: kéo để di chuyển, kéo góc dưới phải để đổi cỡ, chuột phải để xoá / nhân bản.': 'Click '
                                                                                                                                                                'on '
                                                                                                                                                                'the '
                                                                                                                                                                'page '
                                                                                                                                                                'to '
                                                                                                                                                                'place '
                                                                                                                                                                'the '
                                                                                                                                                                'stamp '
                                                                                                                                                                '(or '
                                                                                                                                                                'drag '
                                                                                                                                                                'a '
                                                                                                                                                                'box '
                                                                                                                                                                'to '
                                                                                                                                                                'set '
                                                                                                                                                                'its '
                                                                                                                                                                'size). '
                                                                                                                                                                'Once '
                                                                                                                                                                'placed: '
                                                                                                                                                                'drag '
                                                                                                                                                                'to '
                                                                                                                                                                'move, '
                                                                                                                                                                'drag '
                                                                                                                                                                'the '
                                                                                                                                                                'bottom-right '
                                                                                                                                                                'corner '
                                                                                                                                                                'to '
                                                                                                                                                                'resize, '
                                                                                                                                                                'right-click '
                                                                                                                                                                'to '
                                                                                                                                                                'delete '
                                                                                                                                                                '/ '
                                                                                                                                                                'duplicate.',
 'Bấm vào vị trí cần chèn chữ - nhập nội dung đề xuất': 'Click where text should be inserted - enter the suggestion',
 'Bấm vào vị trí để ghim file (Excel, Word, ảnh...) lên trang': 'Click a spot to pin a file (Excel, Word, image...) '
                                                                'to the page',
 'Bấm vào vị trí để đặt ghi chú': 'Click a position to place a note',
 'Bấm đúp một dòng để tới trang đó': 'Double-click a row to go to its page',
 'Bấm để mở - kéo để di chuyển': 'Click to open - drag to move',
 'Bắt buộc điền': 'Required',
 'Bắt đầu đánh từ trang': 'Start numbering from page',
 'Bị từ chối truy cập khoá (PIN sai hoặc token bị khoá).': 'Key access denied (wrong PIN or token is locked).',
 'Bỏ': 'Remove',
 'Bỏ chọn': 'Deselect',
 'Bỏ chọn vùng\tEsc': 'Deselect region\tEsc',
 "Bỏ chọn: 'An' khớp cả trong 'Anh', 'Toàn'...": "Unticked: 'An' also matches inside 'Anh', 'Toàn'...",
 'Bỏ mật khẩu mở file (cần biết mật khẩu) và gỡ giới hạn in / sao chép / sửa.': 'Remove the open password (password '
                                                                                'required) and lift printing / '
                                                                                'copying / editing restrictions.',
 'Bỏ mọi vùng đánh dấu xoá chưa áp dụng': 'Remove all redaction marks that have not been applied',
 'Bỏ qua (chỉ OCR trang scan)': 'Skip (only OCR scanned pages)',
 'Bỏ qua file không đúng loại:\n': 'Skip files of the wrong type:\n',
 'Bỏ qua nếu lệch dưới': 'Skip if skew is below',
 'Bỏ qua ảnh nhỏ hơn': 'Skip images smaller than',
 'Bỏ đánh dấu xoá': 'Clear redaction marks',
 'Bố cục': 'Layout',
 'Bộ công cụ PDF': 'PDF tools',
 'Bộ công cụ PDF - chạy offline, file không rời khỏi máy': 'PDF toolkit - runs offline, files never leave your '
                                                           'computer',
 'Bộ công cụ PDF nội bộ - chạy offline, file không rời khỏi máy': 'Internal PDF toolkit - runs offline, files never '
                                                                  'leave your computer',
 'Bộ nhớ tạm không có gì để dán': 'Nothing on the clipboard to paste',
 'CHÈN CHỮ: bấm vào vị trí cần chèn, nhập nội dung đề xuất': 'INSERT TEXT: click the insertion point, enter the '
                                                             'suggestion',
 'CHÈN ẢNH: kéo vùng đặt ảnh (hoặc bấm để giữ kích thước gốc)': 'INSERT IMAGE: drag an area to place the image (or '
                                                                'click to keep original size)',
 'CHÈN ẢNH: kéo vùng đặt ảnh (hoặc bấm để giữ kích thước gốc) - kéo ảnh có sẵn để di chuyển': 'INSERT IMAGE: drag an '
                                                                                              'area for the image '
                                                                                              '(or click to keep the '
                                                                                              'original size) - drag '
                                                                                              'an existing image to '
                                                                                              'move it',
 'CHÚ THÍCH MŨI TÊN: kéo từ điểm cần chỉ tới chỗ đặt hộp chữ': 'CALLOUT: drag from the target point to where the '
                                                               'text box goes',
 'CHỌN CHỮ: kéo qua chữ (bấm đúp = cả dòng, Ctrl+A = cả trang) - Ctrl+C sao chép, chuột phải để tô sáng / gạch chân / gạch ngang / tạo liên kết': 'SELECT '
                                                                                                                                                  'TEXT: '
                                                                                                                                                  'drag '
                                                                                                                                                  'over '
                                                                                                                                                  'text '
                                                                                                                                                  '(double-click '
                                                                                                                                                  '= '
                                                                                                                                                  'line, '
                                                                                                                                                  'Ctrl+A '
                                                                                                                                                  '= '
                                                                                                                                                  'page) '
                                                                                                                                                  '- '
                                                                                                                                                  'Ctrl+C '
                                                                                                                                                  'to '
                                                                                                                                                  'copy, '
                                                                                                                                                  'right-click '
                                                                                                                                                  'to '
                                                                                                                                                  'highlight '
                                                                                                                                                  '/ '
                                                                                                                                                  'underline '
                                                                                                                                                  '/ '
                                                                                                                                                  'strike '
                                                                                                                                                  'out '
                                                                                                                                                  '/ '
                                                                                                                                                  'link',
 'CHỌN CHỮ: kéo qua chữ (bấm đúp = cả từ, Shift+bấm = nới vùng chọn, Ctrl+A = cả trang) - Ctrl+C sao chép, chuột phải để tô sáng / gạch chân / gạch ngang / tạo liên kết': 'SELECT '
                                                                                                                                                                           'TEXT: '
                                                                                                                                                                           'drag '
                                                                                                                                                                           'over '
                                                                                                                                                                           'text '
                                                                                                                                                                           '(double-click '
                                                                                                                                                                           '= '
                                                                                                                                                                           'word, '
                                                                                                                                                                           'Shift+click '
                                                                                                                                                                           '= '
                                                                                                                                                                           'extend, '
                                                                                                                                                                           'Ctrl+A '
                                                                                                                                                                           '= '
                                                                                                                                                                           'page) '
                                                                                                                                                                           '- '
                                                                                                                                                                           'Ctrl+C '
                                                                                                                                                                           'to '
                                                                                                                                                                           'copy, '
                                                                                                                                                                           'right-click '
                                                                                                                                                                           'to '
                                                                                                                                                                           'highlight '
                                                                                                                                                                           '/ '
                                                                                                                                                                           'underline '
                                                                                                                                                                           '/ '
                                                                                                                                                                           'strike '
                                                                                                                                                                           'out '
                                                                                                                                                                           '/ '
                                                                                                                                                                           'link',
 'CHỤP VÙNG: kéo vùng cần chụp - ảnh vào bộ nhớ tạm; kéo khung / ô vuông để chỉnh, chuột phải trong khung để lưu ảnh, in, OCR, cắt trang...': 'SNAPSHOT: '
                                                                                                                                              'drag '
                                                                                                                                              'over '
                                                                                                                                              'the '
                                                                                                                                              'area '
                                                                                                                                              'to '
                                                                                                                                              'capture '
                                                                                                                                              '- '
                                                                                                                                              'the '
                                                                                                                                              'image '
                                                                                                                                              'goes '
                                                                                                                                              'to '
                                                                                                                                              'the '
                                                                                                                                              'clipboard; '
                                                                                                                                              'drag '
                                                                                                                                              'the '
                                                                                                                                              'frame '
                                                                                                                                              '/ '
                                                                                                                                              'squares '
                                                                                                                                              'to '
                                                                                                                                              'adjust, '
                                                                                                                                              'right-click '
                                                                                                                                              'inside '
                                                                                                                                              'the '
                                                                                                                                              'frame '
                                                                                                                                              'to '
                                                                                                                                              'save '
                                                                                                                                              'as '
                                                                                                                                              'image, '
                                                                                                                                              'print, '
                                                                                                                                              'OCR, '
                                                                                                                                              'crop '
                                                                                                                                              'the '
                                                                                                                                              'page...',
 'CHỤP VÙNG: kéo vùng cần chụp - ảnh được sao chép vào clipboard (dán vào Word, Outlook...)': 'SNAPSHOT: drag the '
                                                                                              'area to capture - the '
                                                                                              'image is copied to '
                                                                                              'the clipboard (paste '
                                                                                              'into Word, '
                                                                                              'Outlook...)',
 'CHỮ KÝ TAY: kéo vùng đặt chữ ký (hoặc bấm để dùng cỡ mặc định)': 'HANDWRITTEN SIGNATURE: drag an area for the '
                                                                   'signature (or click for the default size)',
 'CHỮ KÝ TAY: kéo vùng đặt chữ ký (hoặc bấm để dùng cỡ mặc định) - kéo chữ ký đã đặt để di chuyển': 'HANDWRITTEN '
                                                                                                    'SIGNATURE: drag '
                                                                                                    'an area for the '
                                                                                                    'signature (or '
                                                                                                    'click for the '
                                                                                                    'default size) - '
                                                                                                    'drag a placed '
                                                                                                    'signature to '
                                                                                                    'move it',
 'CON DẤU: bấm vào trang để đặt (hoặc kéo khung) - kéo con dấu đã đặt để di chuyển, góc dưới phải để đổi cỡ': 'STAMP: '
                                                                                                              'click '
                                                                                                              'on '
                                                                                                              'the '
                                                                                                              'page '
                                                                                                              'to '
                                                                                                              'place '
                                                                                                              '(or '
                                                                                                              'drag '
                                                                                                              'a '
                                                                                                              'box) '
                                                                                                              '- '
                                                                                                              'drag '
                                                                                                              'a '
                                                                                                              'placed '
                                                                                                              'stamp '
                                                                                                              'to '
                                                                                                              'move '
                                                                                                              'it, '
                                                                                                              'bottom-right '
                                                                                                              'corner '
                                                                                                              'to '
                                                                                                              'resize',
 'Cam': 'Orange',
 'Cao (mm)': 'Height (mm)',
 'Che trắng (xoá vùng)': 'Whiteout (erase area)',
 'Che trắng vùng': 'White out region',
 'Chia theo': 'Split by',
 'Chia trang thành poster': 'Split page into a poster',
 'Chia đôi trang': 'Split page in half',
 'Chiều dọc (trái | phải)': 'Vertical (left | right)',
 'Chiều ngang (trên / dưới)': 'Horizontal (top / bottom)',
 'Cho phép chỉnh sửa': 'Allow editing',
 'Cho phép ghi chú / điền form': 'Allow annotations / form filling',
 'Cho phép in': 'Allow printing',
 'Cho phép sao chép chữ / ảnh': 'Allow copying text / images',
 'Chuyển PDF sang Word (.docx) để sửa: giữ đoạn văn, bảng, ảnh, font.': 'Convert PDF to Word (.docx) for editing: '
                                                                        'keeps paragraphs, tables, images, and '
                                                                        'fonts.',
 'Chuyển chữ thành nét vẽ': 'Convert text to outlines',
 'Chuyển chữ, hình vẽ và ảnh sang thang xám (giữ chữ chọn được).': 'Convert text, drawings, and images to grayscale '
                                                                   '(text remains selectable).',
 'Chuyển email đã lưu (.eml) sang PDF: người gửi, người nhận, ngày, nội dung, ảnh trong thư.': 'Convert a saved '
                                                                                               'email (.eml) to PDF: '
                                                                                               'sender, recipient, '
                                                                                               'date, content, and '
                                                                                               'images in the email.',
 'Chuyển file Office sang PDF bằng Microsoft Office trên máy (giữ đúng định dạng nhất).': 'Convert Office files to '
                                                                                          'PDF using Microsoft '
                                                                                          'Office on this computer '
                                                                                          '(most accurate '
                                                                                          'formatting).',
 'Chuyển file TXT, Markdown, HTML, CSV (thành bảng), JSON, XML sang PDF có định dạng.': 'Convert TXT, Markdown, '
                                                                                        'HTML, CSV (as a table), '
                                                                                        'JSON, and XML files to a '
                                                                                        'formatted PDF.',
 'Chuyển giao diện Sáng / Tối (Theo Windows: chọn ở trang chủ)': 'Switch between Light and Dark (Follow Windows: '
                                                                 'choose on the home screen)',
 'Chuyển mỗi trang PDF thành 1 slide PowerPoint (.pptx): chữ sửa được trên nền ảnh, hoặc giữ nguyên dạng ảnh.': 'Turn '
                                                                                                                'each '
                                                                                                                'PDF '
                                                                                                                'page '
                                                                                                                'into '
                                                                                                                'a '
                                                                                                                'PowerPoint '
                                                                                                                'slide '
                                                                                                                '(.pptx): '
                                                                                                                'editable '
                                                                                                                'text '
                                                                                                                'over '
                                                                                                                'the '
                                                                                                                'page '
                                                                                                                'graphics, '
                                                                                                                'or '
                                                                                                                'the '
                                                                                                                'page '
                                                                                                                'as '
                                                                                                                'an '
                                                                                                                'image.',
 'Chuyển sang Excel': 'Convert to Excel',
 'Chuyển sang PowerPoint': 'Convert to PowerPoint',
 'Chuyển sang Word': 'Convert to Word',
 'Chuyển sang ảnh': 'Convert to images',
 'Chuyển sách điện tử EPUB, MOBI, FB2, truyện tranh CBZ, file XPS / SVG sang PDF.': 'Convert EPUB, MOBI, FB2 '
                                                                                    'e-books, CBZ comics, and XPS / '
                                                                                    'SVG files to PDF.',
 'Chuyển thành ảnh': 'Convert to image',
 'Chuyển trang thành ảnh': 'Convert page to image',
 'Chuyển đen trắng': 'Convert to black and white',
 'Chuyển đổi': 'Convert',
 'Chuyển đổi sang PDF': 'Convert to PDF',
 'Chuyển đổi từ PDF': 'Convert from PDF',
 'Chuyển ảnh JPG, PNG, TIFF, WEBP, BMP, PSD... thành PDF. Kéo thả để sắp thứ tự.': 'Convert JPG, PNG, TIFF, WEBP, '
                                                                                   'BMP, PSD... images to PDF. Drag '
                                                                                   'and drop to reorder.',
 'Chuẩn hoá khổ giấy': 'Standardize paper size',
 'Chuẩn khổ giấy': 'Standard paper size',
 'Chèn': 'Insert',
 'Chèn chữ': 'Insert text',
 'Chèn chữ (đề xuất)': 'Insert text (suggestion)',
 'Chèn chữ ký': 'Insert signature',
 'Chèn chữ mờ (MẬT, BẢN NHÁP, tên công ty...) hoặc logo mờ lên các trang.': 'Add a watermark (CONFIDENTIAL, DRAFT, '
                                                                            'company name...) or a faded logo to the '
                                                                            'pages.',
 'Chèn con dấu': 'Insert stamp',
 'Chèn con dấu (stamp)': 'Insert stamp',
 'Chèn con dấu / stamp (ĐÃ DUYỆT, APPROVED, ảnh con dấu scan...) - bấm vào trang để đặt, kéo để di chuyển': 'Insert '
                                                                                                            'a stamp '
                                                                                                            '(APPROVED, '
                                                                                                            'scanned '
                                                                                                            'seal '
                                                                                                            'image...) '
                                                                                                            '- click '
                                                                                                            'the '
                                                                                                            'page to '
                                                                                                            'place, '
                                                                                                            'drag to '
                                                                                                            'move',
 'Chèn file': 'Insert file',
 'Chèn file vào sau trang này...': 'Insert file after this page...',
 'Chèn liên kết': 'Insert link',
 'Chèn liên kết web / email / trang': 'Insert a web link / email / page link',
 'Chèn nhanh': 'Quick insert',
 'Chèn thêm PDF/ảnh vào sau trang đang chọn': 'Insert additional PDF/images after the selected page',
 'Chèn trang mục lục (có số trang, bấm để nhảy tới) từ bookmark của file.': 'Insert a table of contents page (with '
                                                                            'page numbers, clickable to jump) from '
                                                                            "the file's bookmarks.",
 'Chèn trang trắng vào đầu, cuối, sau một trang hoặc xen sau mỗi trang.': 'Insert blank pages at the start, end, '
                                                                          'after a specific page, or interleaved '
                                                                          'after every page.',
 'Chèn trang từ máy scan': 'Insert pages from scanner',
 'Chèn vào sau trang đang xem': 'Insert after the current page',
 'Chèn ảnh': 'Insert image',
 'Chèn ảnh vào vùng...': 'Insert image into region...',
 'Chèn ảnh: kéo vùng đặt ảnh': 'Insert image: drag an area to place it',
 'Chính giữa trang': 'Center on page',
 'Chú thích có mũi tên': 'Callout',
 'Chú thích mũi tên': 'Callout',
 'Chưa chọn chữ nào.': 'No text selected.',
 'Chưa chọn file PDF mẫu.': 'No template PDF file selected.',
 'Chưa chọn file đính kèm.': 'No attachment file selected.',
 'Chưa chọn file ảnh.': 'No image file selected.',
 'Chưa chọn ảnh': 'No image selected',
 'Chưa chọn ảnh con dấu.': 'No stamp image selected.',
 'Chưa cài vào máy (đang chạy trực tiếp file exe)': 'Not installed (running the exe file directly)',
 'Chưa có vùng nào được đánh dấu xoá - dùng Đánh dấu xoá hoặc Tìm & đánh dấu': 'No areas are marked yet - use Mark '
                                                                               'for redaction or Find & mark',
 'Chưa có đường dẫn lưu file.': 'No save path set.',
 'Chưa mở file': 'No file open',
 'Chưa nhập nội dung dấu.': 'No stamp content entered.',
 'Chưa nhập nội dung header / footer.': 'No header / footer content entered.',
 'Chưa nhập nội dung watermark.': 'No watermark content entered.',
 'Chưa đặt mật khẩu mở file và cũng không giới hạn quyền nào.': 'No open password set and no permission restrictions '
                                                                'applied.',
 'Chương trình chính - © 2026 TuanPA': 'Main program - © 2026 TuanPA',
 'Chương trình dùng các thành phần mã nguồn mở dưới đây, mỗi thành phần thuộc bản quyền của tác giả tương ứng và được dùng nguyên bản theo giấy phép ghi kèm.': 'The '
                                                                                                                                                                'program '
                                                                                                                                                                'uses '
                                                                                                                                                                'the '
                                                                                                                                                                'open-source '
                                                                                                                                                                'components '
                                                                                                                                                                'below. '
                                                                                                                                                                'Each '
                                                                                                                                                                'is '
                                                                                                                                                                'copyrighted '
                                                                                                                                                                'by '
                                                                                                                                                                'its '
                                                                                                                                                                'respective '
                                                                                                                                                                'authors '
                                                                                                                                                                'and '
                                                                                                                                                                'used '
                                                                                                                                                                'unmodified '
                                                                                                                                                                'under '
                                                                                                                                                                'the '
                                                                                                                                                                'license '
                                                                                                                                                                'shown.',
 'Chạy lại': 'Run again',
 'Chạy offline - file không rời khỏi máy. Riêng dấu thời gian (TSA) chỉ gửi mã băm của file tới máy chủ dấu thời gian.': 'Runs '
                                                                                                                         'offline '
                                                                                                                         '- '
                                                                                                                         'files '
                                                                                                                         'never '
                                                                                                                         'leave '
                                                                                                                         'your '
                                                                                                                         'computer. '
                                                                                                                         'Only '
                                                                                                                         'the '
                                                                                                                         'timestamp '
                                                                                                                         'feature '
                                                                                                                         '(TSA) '
                                                                                                                         'sends '
                                                                                                                         'a '
                                                                                                                         'hash '
                                                                                                                         'of '
                                                                                                                         'the '
                                                                                                                         'file '
                                                                                                                         'to '
                                                                                                                         'the '
                                                                                                                         'timestamp '
                                                                                                                         'server.',
 'Chạy offline. Khoảng 3-8 giây / trang tuỳ máy.': 'Runs offline. About 3-8 seconds / page depending on the '
                                                   'computer.',
 'Chất lượng JPEG': 'JPEG quality',
 'Chất lượng JPG': 'JPG quality',
 'Chất lượng JPG / WEBP': 'JPG / WEBP quality',
 'Chặn in / sao chép': 'Restrict printing / copying',
 'Chế độ': 'Mode',
 'Chế độ sáng': 'Light mode',
 'Chế độ tối': 'Dark mode',
 'Chế độ xem': 'View mode',
 'Chỉ cài được từ bản exe (PDF Toolkit.exe).': 'Can only be installed from the exe version (PDF Toolkit.exe).',
 'Chỉ cần khi file PDF đầu vào có mật khẩu mở': 'Only needed if the input PDF has an open password',
 'Chỉ hiện dấu ✔ xác nhận đã ký (không chữ, không ảnh)': 'Show only a ✔ mark confirming the signature (no text, no '
                                                         'image)',
 'Chỉ khớp nguyên từ': 'Whole words only',
 'Chỉ kiểm tra tính toàn vẹn (file không bị sửa sau khi ký). Trạng thái thu hồi của chứng thư cần kiểm tra bằng Adobe Acrobat Reader hoặc phần mềm của nhà cung cấp CA.': 'Checks '
                                                                                                                                                                          'integrity '
                                                                                                                                                                          'only '
                                                                                                                                                                          '(file '
                                                                                                                                                                          'not '
                                                                                                                                                                          'modified '
                                                                                                                                                                          'after '
                                                                                                                                                                          'signing). '
                                                                                                                                                                          'Certificate '
                                                                                                                                                                          'revocation '
                                                                                                                                                                          'status '
                                                                                                                                                                          'must '
                                                                                                                                                                          'be '
                                                                                                                                                                          'checked '
                                                                                                                                                                          'using '
                                                                                                                                                                          'Adobe '
                                                                                                                                                                          'Acrobat '
                                                                                                                                                                          'Reader '
                                                                                                                                                                          'or '
                                                                                                                                                                          'the '
                                                                                                                                                                          'CA '
                                                                                                                                                                          "provider's "
                                                                                                                                                                          'software.',
 'Chỉ số dưới': 'Subscript',
 'Chỉ số trên': 'Superscript',
 'Chỉ thu nhỏ trang quá khổ': 'Shrink oversized pages only',
 'Chỉ trang chẵn': 'Even pages only',
 'Chỉ trang lẻ': 'Odd pages only',
 'Chỉ trang đang xem': 'Current page only',
 'Chỉ tối ưu cấu trúc file, không giảm chất lượng ảnh.': 'Only optimizes file structure, without reducing image '
                                                         'quality.',
 'Chỉ xoay': 'Rotate only',
 'Chỉ xoá nền trắng': 'Only remove the white background',
 'Chỉnh sáng / tương phản': 'Brightness / contrast',
 'Chỉnh sửa': 'Edit',
 'Chỉnh sửa & Ký số': 'Edit & Digital Signature',
 'Chỉnh độ sáng, tương phản, độ bão hoà, gamma, làm nét - làm rõ bản scan mờ. Trang sẽ thành ảnh.': 'Adjust '
                                                                                                    'brightness, '
                                                                                                    'contrast, '
                                                                                                    'saturation, '
                                                                                                    'gamma and '
                                                                                                    'sharpness to '
                                                                                                    'clean up faint '
                                                                                                    'scans. Pages '
                                                                                                    'become images.',
 'Chọn chữ': 'Select text',
 'Chọn file': 'Select file',
 'Chọn file PDF chứa trang thay thế': 'Choose the PDF with the replacement pages',
 'Chọn file để ghim lên trang': 'Choose a file to pin to the page',
 'Chọn file để gộp': 'Select files to merge',
 'Chọn file để so sánh với file đang mở': 'Choose a file to compare with the open file',
 'Chọn hoặc dán ảnh có con dấu (ảnh scan, ảnh chụp văn bản đã đóng dấu...).': 'Choose or paste an image with a stamp '
                                                                              '(scan, photo of a stamped '
                                                                              'document...).',
 'Chọn hết chữ trang': 'Select all text on page',
 'Chọn hết chữ trang  (Ctrl+A)': 'Select all text on page  (Ctrl+A)',
 'Chọn màu': 'Choose color',
 'Chọn màu chữ': 'Choose text color',
 'Chọn màu nền': 'Choose background color',
 'Chọn thư mục': 'Select folder',
 'Chọn toàn bộ chữ trong trang': 'Select all text on page',
 'Chọn ảnh chữ ký': 'Choose signature image',
 'Chọn ảnh chữ ký tay': 'Select handwritten signature image',
 'Chọn ảnh con dấu': 'Choose a stamp image',
 'Chọn ảnh để chèn': 'Select image to insert',
 'Chọn ảnh...': 'Select image...',
 'Chọn...': 'Select...',
 'Chống sao chép': 'Copy protection',
 'Chồng trang (letterhead / mẫu nền)': 'Overlay page (letterhead / background template)',
 'Chồng trang mẫu': 'Overlay template page',
 'Chụp vùng': 'Snapshot',
 'Chủ yếu tiếng Anh, có ít tiếng Việt': 'Mostly English, with some Vietnamese',
 'Chủ đề': 'Theme',
 'Chức năng cài đặt chỉ dùng được trong bản exe (PDF Toolkit.exe).': 'Installing is only available in the exe '
                                                                     'version (PDF Toolkit.exe).',
 'Chứng thư dùng khoá ECC — phiên bản này chỉ hỗ trợ khoá RSA (đa số token tại VN).': 'Certificate uses an ECC key — '
                                                                                      'this version only supports '
                                                                                      'RSA keys (used by most tokens '
                                                                                      'in Vietnam).',
 'Chứng thư số:': 'Digital certificate:',
 'Chừa lề 1 cm': '1 cm margin',
 'Chừa lề 1 cm khi căn sát mép': 'Keep a 1 cm margin from the page edge',
 'Chừa lề quanh nội dung': 'Leave a margin around the content',
 'Chữ': 'Text',
 'Chữ cần xoá (mỗi dòng 1 cụm)': 'Text to redact (one phrase per line)',
 'Chữ cần xoá:': 'Text to redact:',
 'Chữ hiển thị (để trống = biến vùng đã kéo trên chữ / ảnh có sẵn thành liên kết):': 'Display text (leave empty to '
                                                                                     'turn the dragged area over '
                                                                                     'existing text / image into a '
                                                                                     'link):',
 'Chữ hiển thị trống.': 'The display text is empty.',
 'Chữ hiển thị:': 'Display text:',
 'Chữ ký & con dấu': 'Signature & stamp',
 'Chữ ký số trong file': 'Digital signatures in the file',
 'Chữ ký tay': 'Handwritten signature',
 'Chữ ký đã lưu': 'Saved signature',
 'Chữ nhận dạng được (OCR)': 'Recognized text (OCR)',
 'Chữ sửa được (nền = hình ảnh / hình vẽ của trang)': 'Editable text (background = page images / drawings)',
 'Chữ vẫn giữ nguyên (chọn, tìm kiếm được). Có thể hoàn tác (Ctrl+Z).': 'Text remains unchanged (selectable, '
                                                                        'searchable). Can be undone (Ctrl+Z).',
 'Chữ ẩn (không nhìn thấy)': 'Hidden text (invisible)',
 'Chữ → nét vẽ': 'Text → outlines',
 'Clipboard không có ảnh. Hãy chụp vùng con dấu (Win+Shift+S) hoặc copy ảnh rồi bấm lại.': 'The clipboard has no '
                                                                                           'image. Capture the stamp '
                                                                                           'area (Win+Shift+S) or '
                                                                                           'copy an image, then try '
                                                                                           'again.',
 'Co chữ cho vừa nếu nội dung mới dài hơn': 'Shrink text to fit if the new content is longer',
 'Co chữ cho vừa độ rộng dòng cũ nếu nội dung mới dài hơn': 'Shrink text to fit the old line width if the new '
                                                            'content is longer',
 'Con dấu': 'Stamp',
 'Con dấu (stamp) - nằm trên trang như ghi chú: ẩn / xoá / cố định được, không trộn vào nội dung': 'Stamp - sits on '
                                                                                                   'the page like a '
                                                                                                   'comment: can be '
                                                                                                   'hidden / deleted '
                                                                                                   '/ flattened, not '
                                                                                                   'merged into the '
                                                                                                   'content',
 'Con dấu chữ': 'Text stamp',
 'Con dấu tròn': 'Round seal',
 'Cuối file': 'End of file',
 'Cuối trang - giữa': 'Bottom of page - center',
 'Cuối trang - phải': 'Bottom of page - right',
 'Cuối trang - trái': 'Bottom of page - left',
 'Cuộn phải - còn nút phía sau': 'Scroll right - more buttons after',
 'Cuộn trái - còn nút phía trước': 'Scroll left - more buttons before',
 'Cài PDF Toolkit cho tài khoản Windows này - <b>không cần quyền administrator</b>.': 'Install PDF Toolkit for this '
                                                                                      'Windows account - <b>no '
                                                                                      'administrator rights '
                                                                                      'needed</b>.',
 'Cài PDF Toolkit cho tài khoản này (không cần quyền admin), add-in Office': 'Install PDF Toolkit for this account '
                                                                             '(no admin rights needed), Office '
                                                                             'add-in',
 'Cài đặt': 'Install',
 'Cài đặt / cập nhật vào máy...': 'Install / update on this PC...',
 'Cài đặt PDF Toolkit': 'Install PDF Toolkit',
 'Cài đặt vào máy': 'Install on this PC',
 'Các chỗ tìm thấy được đánh dấu khung đỏ để xem lại; bấm <b>Xoá vĩnh viễn</b> để xoá hẳn.': 'Matches are marked '
                                                                                             'with a red frame for '
                                                                                             'review; click '
                                                                                             '<b>Redact '
                                                                                             'permanently</b> to '
                                                                                             'remove them.',
 'Các file hỗ trợ (': 'Supported files (',
 'Các khoảng trang': 'Page ranges',
 'Các lớp': 'Layers',
 'Các lựa chọn:': 'Choices:',
 'Các trang đang chọn': 'Selected pages',
 'Các ô form': 'Form fields',
 'Cách cắt': 'Cut method',
 'Cách mép giấy': 'Distance from paper edge',
 'Cách tách': 'Split method',
 'Còn sót (đang được dùng): ': 'Left behind (in use): ',
 'Có thêm nội dung sau khi ký (vd: chữ ký khác)': 'Additional content added after signing (e.g. another signature)',
 'Công cụ': 'Tool',
 'Công cụ này không có tuỳ chọn.': 'This tool has no options.',
 'Công cụ phổ biến': 'Popular tools',
 'Công cụ đang chạy - bấm Dừng trước.': 'A tool is running - click Stop first.',
 'Công cụ đang chạy. Dừng và thoát?': 'A tool is running. Stop and exit?',
 'Công dụng': 'Purpose',
 'Căn chỉnh': 'Align',
 'Căn chỉnh trên trang': 'Align on page',
 'Căn dưới': 'Align bottom',
 'Căn giữa': 'Center',
 'Căn giữa (dọc)': 'Center vertically',
 'Căn giữa (ngang)': 'Center horizontally',
 'Căn giữa trang trên giấy': 'Center the page on the paper',
 'Căn phải': 'Align right',
 'Căn trái': 'Align left',
 'Căn trên': 'Align top',
 'Căn vật đang chọn sát lề dưới trang': 'Align the selected object to the bottom of the page',
 'Căn vật đang chọn sát lề phải trang': 'Align the selected object to the right of the page',
 'Căn vật đang chọn sát lề trái trang': 'Align the selected object to the left of the page',
 'Căn vật đang chọn sát lề trên trang': 'Align the selected object to the top of the page',
 'Căn vật đang chọn vào giữa trang theo chiều dọc': 'Center the selected object vertically on the page',
 'Căn vật đang chọn vào giữa trang theo chiều ngang': 'Center the selected object horizontally on the page',
 'Căn đều hai bên': 'Justify',
 'Cần Internet: chỉ gửi mã băm (SHA-256) của file tới máy chủ, không gửi nội dung file. Không sửa file sau khi gắn dấu.': 'Requires '
                                                                                                                          'Internet: '
                                                                                                                          'only '
                                                                                                                          'the '
                                                                                                                          "file's "
                                                                                                                          'SHA-256 '
                                                                                                                          'hash '
                                                                                                                          'is '
                                                                                                                          'sent '
                                                                                                                          'to '
                                                                                                                          'the '
                                                                                                                          'server, '
                                                                                                                          'never '
                                                                                                                          'its '
                                                                                                                          'content. '
                                                                                                                          'Do '
                                                                                                                          'not '
                                                                                                                          'edit '
                                                                                                                          'the '
                                                                                                                          'file '
                                                                                                                          'after '
                                                                                                                          'timestamping.',
 'Cần Microsoft Office trên máy. File ODT/ODS/ODP, RTF, CSV cũng được.': 'Requires Microsoft Office on this '
                                                                         'computer. ODT/ODS/ODP, RTF, and CSV files '
                                                                         'are also supported.',
 'Cần hỗ trợ?': 'Need help?',
 'Cần hỗ trợ? Liên hệ bộ phận IT': 'Need help? Contact the IT team',
 'Cần hỗ trợ? Liên hệ qua email': 'Need help? Contact us by email',
 'Cầu nối Python - Qt': 'Python - Qt bindings',
 'Cắt': 'Cut',
 'Cắt bớt lề theo số mm, hoặc tự cắt sát nội dung (bỏ khoảng trắng thừa).': 'Trim margins by a number of mm, or '
                                                                            'auto-crop to the content (remove extra '
                                                                            'white space).',
 'Cắt dưới': 'Crop bottom',
 'Cắt lề': 'Crop margins',
 'Cắt lề trang': 'Crop page margins',
 'Cắt mỗi trang thành 2 (hoặc nhiều) trang - dùng cho bản scan 2 trang sách trên 1 mặt.': 'Split each page into 2 '
                                                                                          '(or more) pages - for '
                                                                                          'scans of 2 book pages on '
                                                                                          'one side.',
 'Cắt phải': 'Crop right',
 'Cắt sát con dấu': 'Trim tightly around the stamp',
 'Cắt trang theo vùng': 'Crop page to region',
 'Cắt trái': 'Crop left',
 'Cắt trên': 'Crop top',
 'Cắt vật đang chọn trên trang': 'Cut the selected object on the page',
 'Cố định form': 'Flatten forms',
 'Cố định ghi chú': 'Flatten comments',
 'Cố định ghi chú & form vào trang': 'Flatten comments & forms',
 'Cỡ chữ': 'Font size',
 'Cỡ chữ (sách)': 'Font size (book)',
 'Cỡ chữ:': 'Font size:',
 'Danh sách chọn': 'List box',
 'Danh sách ghi chú': 'Comment list',
 'Danh sách thả xuống': 'Drop-down list',
 'Driver token chưa cài đặt / không tải được.': 'Token driver not installed / could not be loaded.',
 'Dung lượng tối đa': 'Maximum size',
 'Dán': 'Paste',
 'Dán (Ctrl+V) vào Word, Outlook, Zalo...': 'Paste (Ctrl+V) into Word, Outlook, Zalo...',
 'Dán vật đã cắt / sao chép, hoặc ảnh / chữ từ bộ nhớ tạm, vào trang': 'Paste the cut / copied object, or an image / '
                                                                       'text from the clipboard, onto the page',
 'Dán ảnh (Ctrl+V)': 'Paste image (Ctrl+V)',
 'Dán ảnh vừa chụp màn hình (Win+Shift+S) hoặc copy từ Word / trình duyệt': 'Paste a screenshot (Win+Shift+S) or an '
                                                                            'image copied from Word / a browser',
 'Dòng phụ': 'Subline',
 'Dùng con dấu chữ': 'Use a text stamp',
 'Dùng hộp thoại chọn file của Windows': 'Use the Windows file dialog',
 'Dùng khi gửi file in ấn hoặc không muốn người nhận chép / sửa chữ. Liên kết sẽ bị bỏ.': 'Use for print-ready files '
                                                                                          'or when recipients should '
                                                                                          'not copy / edit the text. '
                                                                                          'Links are removed.',
 'Dùng máy scan có driver WIA của Windows (hầu hết máy scan / máy in đa năng). Sau khi quét có thể bấm <b>OCR</b> để tìm và copy được chữ.': 'Works '
                                                                                                                                             'with '
                                                                                                                                             'scanners '
                                                                                                                                             'that '
                                                                                                                                             'have '
                                                                                                                                             'a '
                                                                                                                                             'Windows '
                                                                                                                                             'WIA '
                                                                                                                                             'driver '
                                                                                                                                             '(most '
                                                                                                                                             'scanners '
                                                                                                                                             '/ '
                                                                                                                                             'multifunction '
                                                                                                                                             'printers). '
                                                                                                                                             'After '
                                                                                                                                             'scanning, '
                                                                                                                                             'click '
                                                                                                                                             '<b>OCR</b> '
                                                                                                                                             'to '
                                                                                                                                             'make '
                                                                                                                                             'the '
                                                                                                                                             'text '
                                                                                                                                             'searchable '
                                                                                                                                             'and '
                                                                                                                                             'copyable.',
 'Dùng trang 1 của mẫu cho mọi trang': 'Use page 1 of the template for all pages',
 'Dùng để gỡ giới hạn sau này. Để trống = tự sinh ngẫu nhiên': 'Used to remove restrictions later. Leave blank to '
                                                               'auto-generate randomly',
 'Dùng ảnh chữ ký...': 'Use signature image...',
 'Dùng ảnh con dấu (scan)...': 'Use a stamp image (scan)...',
 'Dưới - giữa': 'Bottom - center',
 'Dưới - phải': 'Bottom - right',
 'Dưới - trái': 'Bottom - left',
 'Dấu': 'Stamp',
 'Dấu thời gian': 'Timestamp',
 'Dấu thời gian: ': 'Timestamp: ',
 'Dọc': 'Portrait',
 'Dọc trước (trên→dưới)': 'Vertical first (top→bottom)',
 'Dừng': 'Stop',
 'Dữ liệu': 'Data',
 'Dựng lại tất cả các trang': 'Rebuild all pages',
 "Email trong Outlook: kéo thư ra Desktop sẽ được file .msg; dùng 'Lưu thành' > .eml hoặc Mail Toolkit để xuất.": 'Email '
                                                                                                                  'in '
                                                                                                                  'Outlook: '
                                                                                                                  'dragging '
                                                                                                                  'a '
                                                                                                                  'message '
                                                                                                                  'to '
                                                                                                                  'the '
                                                                                                                  'Desktop '
                                                                                                                  'produces '
                                                                                                                  'a '
                                                                                                                  '.msg '
                                                                                                                  'file; '
                                                                                                                  'use '
                                                                                                                  "'Save "
                                                                                                                  "As' "
                                                                                                                  '> '
                                                                                                                  '.eml '
                                                                                                                  'or '
                                                                                                                  'Mail '
                                                                                                                  'Toolkit '
                                                                                                                  'to '
                                                                                                                  'export.',
 'Excel: co vừa chiều ngang 1 trang giấy': 'Excel: fit width to 1 page',
 'Excel: xuất tất cả sheet': 'Excel: export all sheets',
 'Explorer: ảnh thu nhỏ trang đầu và khung xem trước (Preview pane) cho file PDF': 'Explorer: first-page thumbnails '
                                                                                   'and Preview pane for PDF files',
 'File 1 = bản cũ, file 2 = bản mới.': 'File 1 = old version, file 2 = new version.',
 'File 1 = mặt trước, file 2 = mặt sau.': 'File 1 = front side, file 2 = back side.',
 'File PDF mẫu': 'Template PDF file',
 'File có mật khẩu': 'File has a password',
 'File gốc': 'Original file',
 'File không có ô form, hoặc chưa nhập giá trị.': 'The file has no form fields, or no values have been entered.',
 'File nguồn:': 'Source files:',
 'File nhúng / đính kèm': 'Embedded / attached files',
 'File này có chữ ký số. Lưu thay đổi sẽ làm các chữ ký hiện có MẤT HIỆU LỰC.\n\nVẫn lưu?': 'This file has digital '
                                                                                            'signatures. Saving '
                                                                                            'changes will INVALIDATE '
                                                                                            'the existing '
                                                                                            'signatures.\n'
                                                                                            '\n'
                                                                                            'Save anyway?',
 'File này không có chữ ký số.': 'This file has no digital signature.',
 'File scan có nửa đầu là mặt trước, nửa sau là mặt sau → sắp lại đúng thứ tự 1, 2, 3, 4...': 'A scan whose first '
                                                                                              'half holds the front '
                                                                                              'sides and second half '
                                                                                              'the back sides → '
                                                                                              'reorder to 1, 2, 3, '
                                                                                              '4...',
 'File thứ 2 có thứ tự ngược (scan lật cả xấp)': 'The 2nd file is in reverse order (scanned as a flipped stack)',
 'File zip (*.zip)': 'Zip file (*.zip)',
 'File đã có chữ ký số và đã bị chỉnh sửa trong phiên này - các chữ ký cũ sẽ mất hiệu lực. Vẫn ký?': 'The file '
                                                                                                     'already has '
                                                                                                     'digital '
                                                                                                     'signatures and '
                                                                                                     'has been '
                                                                                                     'edited in this '
                                                                                                     'session - the '
                                                                                                     'old signatures '
                                                                                                     'will become '
                                                                                                     'invalid. Sign '
                                                                                                     'anyway?',
 'File đã thay đổi. Lưu lại trước khi tiếp tục?': 'The file has changed. Save before continuing?',
 'File đã tồn tại. Ghi đè?': 'The file already exists. Overwrite?',
 'File đính kèm': 'Attachment',
 'File ảnh': 'Image file',
 'Font chữ': 'Font',
 'Font chữ (chỉ liệt kê font có đủ dấu tiếng Việt)': 'Font (only fonts with full Vietnamese diacritics are listed)',
 'Font chữ Windows - nhóm đầu có đủ dấu tiếng Việt': 'Windows fonts - the first group fully supports Vietnamese',
 'Font chữ khi sửa chữ': 'Fonts for text editing',
 'Font lớp chữ ẩn sau khi OCR': 'Font for the hidden OCR text layer',
 'Font này thiếu một số chữ có dấu tiếng Việt - chữ thiếu sẽ hiện bằng font khác': 'This font lacks some Vietnamese '
                                                                                   'accented letters - missing '
                                                                                   'letters will use another font',
 'Font thiếu dấu tiếng Việt:': 'Fonts missing Vietnamese accents:',
 'Form & chữ ký': 'Forms & signature',
 'GHI CHÚ: bấm vào vị trí đặt ghi chú': 'NOTE: click a position to place the note',
 'GNU AGPL v3 - phần mềm tự do: được dùng, sao chép, sửa đổi và phân phối lại theo giấy phép. KHÔNG CÓ BẢO HÀNH. Mã nguồn đi kèm chương trình (tab Giấy phép).': 'GNU '
                                                                                                                                                                 'AGPL '
                                                                                                                                                                 'v3 '
                                                                                                                                                                 '- '
                                                                                                                                                                 'free '
                                                                                                                                                                 'software: '
                                                                                                                                                                 'you '
                                                                                                                                                                 'may '
                                                                                                                                                                 'use, '
                                                                                                                                                                 'copy, '
                                                                                                                                                                 'modify '
                                                                                                                                                                 'and '
                                                                                                                                                                 'redistribute '
                                                                                                                                                                 'it '
                                                                                                                                                                 'under '
                                                                                                                                                                 'the '
                                                                                                                                                                 'license. '
                                                                                                                                                                 'NO '
                                                                                                                                                                 'WARRANTY. '
                                                                                                                                                                 'The '
                                                                                                                                                                 'source '
                                                                                                                                                                 'code '
                                                                                                                                                                 'is '
                                                                                                                                                                 'bundled '
                                                                                                                                                                 'with '
                                                                                                                                                                 'the '
                                                                                                                                                                 'program '
                                                                                                                                                                 '(Licenses '
                                                                                                                                                                 'tab).',
 'Ghi chú': 'Note',
 'Ghi chú & form': 'Notes & forms',
 'Ghi chú / bình luận': 'Notes / comments',
 'Ghi chú / hình vẽ': 'Notes / drawings',
 'Ghi chú / liên kết trên trang được xoay sẽ bị bỏ.': 'Notes / links on rotated pages are removed.',
 'Ghi dấu ngắt trang': 'Mark page breaks',
 'Ghép 1 trang dài': 'Combine into one long page',
 'Ghép file mặt trước (trang lẻ) và file mặt sau (trang chẵn) khi scan 2 mặt bằng máy 1 mặt.': 'Merge the front-side '
                                                                                               'file (odd pages) and '
                                                                                               'back-side file (even '
                                                                                               'pages) when scanning '
                                                                                               'both sides with a '
                                                                                               'single-sided '
                                                                                               'scanner.',
 'Ghép thành 1 trang dài': 'Combine into one long page',
 'Giao diện': 'User interface',
 'Giao diện: Sáng / Tối / Theo Windows': 'Appearance: Light / Dark / Follow Windows',
 'Giá trị': 'Value',
 'Giá trị mặc định:': 'Default value:',
 'Giả lập bản scan': 'Simulate a scan',
 'Giả lập scan': 'Simulate scan',
 'Giảm dung lượng: chọn mức nén, hoặc nhập dung lượng mong muốn (vd dưới 1 MB để gửi email).': 'Reduce file size: '
                                                                                               'choose a compression '
                                                                                               'level, or enter the '
                                                                                               'desired size (e.g. '
                                                                                               'under 1 MB for '
                                                                                               'email).',
 'Giấy phép': 'License',
 'Giấy phép bên thứ ba': 'Third-party licenses',
 'Giữ hình ảnh của bản scan, xoá chữ cũ (logo, con dấu, bảng giữ nguyên)': "Keep the scan's graphics, remove the old "
                                                                           'text (logos, stamps, tables kept)',
 'Giữ hướng từng trang': "Keep each page's orientation",
 'Giữ liên kết': 'Keep links',
 'Giữ màu': 'Keep color',
 'Giữ nguyên (chỉ dựng lại trang scan)': 'Keep as is (rebuild scanned pages only)',
 'Giữ nguyên ảnh': 'Keep the image as is',
 'Giữ ô form': 'Keep form fields',
 'Giữa dọc': 'Middle',
 'Giữa trang': 'Center of page',
 'Gop_PDF': 'Merged_PDF',
 'Góc nghiêng': 'Tilt angle',
 'Góc xoay': 'Rotation angle',
 'Gõ chữ nhanh': 'Typewriter',
 'GẠCH CHÂN: kéo qua đoạn chữ': 'UNDERLINE: drag over the text',
 'GẠCH NGANG: kéo qua đoạn chữ cần đề xuất xoá': 'STRIKEOUT: drag over the text to suggest deleting',
 'Gạch chân': 'Underline',
 'Gạch ngang': 'Strikethrough',
 'Gạch sóng': 'Squiggly',
 'Gần đây:': 'Recent:',
 'Gắn dấu thời gian (TSA)': 'Add timestamp (TSA)',
 'Gắn dấu thời gian tin cậy (RFC 3161) chứng minh file tồn tại, không bị sửa từ thời điểm này. Không cần USB Token.': 'Add '
                                                                                                                      'a '
                                                                                                                      'trusted '
                                                                                                                      'RFC '
                                                                                                                      '3161 '
                                                                                                                      'timestamp '
                                                                                                                      'proving '
                                                                                                                      'the '
                                                                                                                      'file '
                                                                                                                      'existed '
                                                                                                                      'unmodified '
                                                                                                                      'from '
                                                                                                                      'this '
                                                                                                                      'moment. '
                                                                                                                      'No '
                                                                                                                      'USB '
                                                                                                                      'token '
                                                                                                                      'needed.',
 'Gặp lỗi, cần thêm tính năng hoặc cần hướng dẫn sử dụng, vui lòng liên hệ bộ phận IT qua email:': 'For bugs, '
                                                                                                   'feature requests '
                                                                                                   'or help using '
                                                                                                   'the app, please '
                                                                                                   'contact the IT '
                                                                                                   'team by email:',
 'Gặp lỗi, cần thêm tính năng hoặc cần hướng dẫn sử dụng, vui lòng liên hệ qua email:': 'For bugs, feature requests '
                                                                                        'or help using the app, '
                                                                                        'please contact us by email:',
 'Gọi Windows / Office': 'Windows / Office integration',
 'Gộp PDF': 'Merge PDF',
 'Gộp các file đã chọn thành 1 PDF': 'Merge selected files into 1 PDF',
 'Gộp file': 'Merge files',
 'Gộp nhiều file PDF / ảnh thành 1 file. Kéo thả để sắp thứ tự.': 'Merge multiple PDF / image files into one file. '
                                                                  'Drag and drop to reorder.',
 'Gộp nhiều trang / 1 tờ (N-up)': 'Merge multiple pages per sheet (N-up)',
 'Gộp tất cả ảnh vào 1 file': 'Merge all images into 1 file',
 'Gộp xen kẽ (scan 2 mặt)': 'Interleave merge (duplex scan)',
 'Gỡ PDF Toolkit khỏi máy?\nSẽ xoá shortcut, liên kết file PDF, add-in Office và thư mục cài đặt.\nFile PDF của bạn không bị ảnh hưởng.': 'Uninstall '
                                                                                                                                          'PDF '
                                                                                                                                          'Toolkit '
                                                                                                                                          'from '
                                                                                                                                          'this '
                                                                                                                                          'PC?\n'
                                                                                                                                          'Shortcuts, '
                                                                                                                                          'the '
                                                                                                                                          'PDF '
                                                                                                                                          'file '
                                                                                                                                          'association, '
                                                                                                                                          'the '
                                                                                                                                          'Office '
                                                                                                                                          'add-in '
                                                                                                                                          'and '
                                                                                                                                          'the '
                                                                                                                                          'install '
                                                                                                                                          'folder '
                                                                                                                                          'will '
                                                                                                                                          'be '
                                                                                                                                          'removed.\n'
                                                                                                                                          'Your '
                                                                                                                                          'PDF '
                                                                                                                                          'files '
                                                                                                                                          'are '
                                                                                                                                          'not '
                                                                                                                                          'affected.',
 'Gỡ PDF Toolkit khỏi máy?\nSẽ xoá shortcut, liên kết file PDF, add-in Office, thư mục cài đặt và toàn bộ dữ liệu của app (file gần đây, cài đặt, thư viện con dấu).\nFile PDF của bạn không bị ảnh hưởng.': 'Uninstall '
                                                                                                                                                                                                             'PDF '
                                                                                                                                                                                                             'Toolkit?\n'
                                                                                                                                                                                                             'This '
                                                                                                                                                                                                             'removes '
                                                                                                                                                                                                             'shortcuts, '
                                                                                                                                                                                                             'PDF '
                                                                                                                                                                                                             'file '
                                                                                                                                                                                                             'associations, '
                                                                                                                                                                                                             'the '
                                                                                                                                                                                                             'Office '
                                                                                                                                                                                                             'add-in, '
                                                                                                                                                                                                             'the '
                                                                                                                                                                                                             'install '
                                                                                                                                                                                                             'folder '
                                                                                                                                                                                                             'and '
                                                                                                                                                                                                             'all '
                                                                                                                                                                                                             'app '
                                                                                                                                                                                                             'data '
                                                                                                                                                                                                             '(recent '
                                                                                                                                                                                                             'files, '
                                                                                                                                                                                                             'settings, '
                                                                                                                                                                                                             'stamp '
                                                                                                                                                                                                             'library).\n'
                                                                                                                                                                                                             'Your '
                                                                                                                                                                                                             'PDF '
                                                                                                                                                                                                             'files '
                                                                                                                                                                                                             'are '
                                                                                                                                                                                                             'not '
                                                                                                                                                                                                             'affected.',
 'Gỡ cài đặt...': 'Uninstall...',
 'Gỡ mật khẩu': 'Remove password',
 'Gửi email': 'Email',
 'Gửi email hỗ trợ': 'Email support',
 'HUỶ': 'CANCEL',
 'Hai file giống nhau về nội dung chữ.': 'The two files have identical text content.',
 'Hiện / thu gọn bảng ảnh thu nhỏ các trang bên trái': 'Show / collapse the page thumbnail panel on the left',
 'Hiện thanh công cụ': 'Show toolbar',
 'Hiện thông tin người ký (tên, đơn vị, ngày giờ ký)': 'Show signer information (name, organization, signing '
                                                       'date/time)',
 'Hiện tất cả các lớp': 'Show all layers',
 'Hiện tới cấp': 'Show up to level',
 'Hiệu ứng scan': 'Scan effect',
 'Hoàn tác': 'Undo',
 'Hoàn tất.': 'Done.',
 'Hoán đổi 2 trang': 'Swap 2 pages',
 'Hoặc kéo thả file PDF / ảnh vào cửa sổ này.': 'Or drag & drop PDF / image files onto this window.',
 'Hoặc kéo thả file PDF / ảnh vào cửa sổ này.  Phím tắt: Ctrl+O': 'Or drag and drop PDF / image files into this '
                                                                  'window.  Shortcut: Ctrl+O',
 'Huỷ': 'Cancel',
 'HÌNH CHỮ NHẬT: kéo để vẽ': 'RECTANGLE: drag to draw',
 'HÌNH CHỮ NHẬT: kéo để vẽ - kéo viền hình đã vẽ để di chuyển, ô góc dưới phải để đổi cỡ': 'RECTANGLE: drag to draw '
                                                                                           '- drag the edge of a '
                                                                                           'drawn shape to move it, '
                                                                                           'the bottom-right square '
                                                                                           'to resize',
 'HÌNH ELIP: kéo để vẽ': 'ELLIPSE: drag to draw',
 'HÌNH ELIP: kéo để vẽ - kéo viền hình đã vẽ để di chuyển, ô góc dưới phải để đổi cỡ': 'ELLIPSE: drag to draw - drag '
                                                                                       'the edge of a drawn shape to '
                                                                                       'move it, the bottom-right '
                                                                                       'square to resize',
 'Hãy bỏ chọn ít nhất một quyền cần chặn (in, sao chép, chỉnh sửa...).': 'Untick at least one permission to block '
                                                                         '(printing, copying, editing...).',
 'Hãy chọn thư mục lưu kết quả.': 'Please select a folder to save the results.',
 'Hãy chọn đúng 2 trang ở danh sách trang bên trái (giữ Ctrl rồi bấm) để hoán đổi.': 'Select exactly 2 pages in the '
                                                                                     'page list on the left (hold '
                                                                                     'Ctrl and click) to swap.',
 'Hãy cắm USB Token, cài driver / phần mềm quản lý token của nhà cung cấp (VNPT-CA, Viettel-CA, BKAV-CA, FPT-CA...) rồi bấm Làm mới.': 'Please '
                                                                                                                                       'plug '
                                                                                                                                       'in '
                                                                                                                                       'the '
                                                                                                                                       'USB '
                                                                                                                                       'Token, '
                                                                                                                                       'install '
                                                                                                                                       'the '
                                                                                                                                       "provider's "
                                                                                                                                       'driver '
                                                                                                                                       '/ '
                                                                                                                                       'token '
                                                                                                                                       'management '
                                                                                                                                       'software '
                                                                                                                                       '(VNPT-CA, '
                                                                                                                                       'Viettel-CA, '
                                                                                                                                       'BKAV-CA, '
                                                                                                                                       'FPT-CA...), '
                                                                                                                                       'then '
                                                                                                                                       'click '
                                                                                                                                       'Refresh.',
 'Hãy ký vào ô trắng hoặc chọn ảnh chữ ký.': 'Sign in the white box or choose a signature image.',
 'Hãy nhập chữ hiển thị cho liên kết.': 'Please enter display text for the link.',
 'Hãy nhập các trang cần lấy.': 'Please enter the pages to extract.',
 'Hãy nhập nội dung.': 'Please enter some text.',
 'Hãy nhập ít nhất một lựa chọn.': 'Enter at least one choice.',
 'Hãy nhập địa chỉ liên kết.': 'Please enter the link address.',
 'Hãy thêm file cần xử lý.': 'Please add files to process.',
 'Hình chữ nhật': 'Rectangle',
 'Hình elip': 'Ellipse',
 'Hình vẽ': 'Drawing',
 'Hướng': 'Orientation',
 'Hướng dọc': 'Portrait orientation',
 'Hướng ngang': 'Landscape orientation',
 'Hướng nối': 'Direction',
 'Hướng trang': 'Page orientation',
 'Hậu tố': 'Suffix',
 'Hồng': 'Pink',
 'Hỗ trợ': 'Support',
 'HỘP CHỮ: kéo khung (hoặc bấm) để đặt hộp chữ - kéo viền hộp đã đặt để di chuyển': 'TEXT BOX: drag a frame (or '
                                                                                    'click) to place a text box - '
                                                                                    "drag a box's edge to move it",
 'Hộp chữ': 'Text box',
 'In': 'Print',
 'In 1 mặt': 'Print one sided',
 'In 2 mặt - lật theo cạnh dài': 'Print on both sides - flip on long edge',
 'In 2 mặt - lật theo cạnh ngắn': 'Print on both sides - flip on short edge',
 'In 2 mặt, lật theo cạnh ngắn, rồi gập đôi.': 'Print double-sided, flip on short edge, then fold in half.',
 'In 2, 4, 6, 9, 16 trang trên một tờ giấy để tiết kiệm giấy.': 'Print 2, 4, 6, 9, or 16 pages per sheet to save '
                                                                'paper.',
 'In khổ lớn ghép nhiều tờ (poster)': 'Poster (tile across sheets)',
 'In kiểu sách gập (booklet)': 'Booklet (fold-and-bind printing)',
 'In màu': 'Color',
 'In nghiêng': 'Italic',
 'In nhiều trang trên 1 tờ (N-up)': 'Print multiple pages per sheet (N-up)',
 'In poster': 'Poster',
 'In tài liệu': 'Print the document',
 'In vùng chọn': 'Print region',
 'In vùng chọn...': 'Print region...',
 'In đen trắng': 'Black and white',
 'In đậm': 'Bold',
 'KHẨN': 'URGENT',
 'Khay nạp giấy tự động - ADF (quét đến hết giấy)': 'Automatic document feeder - ADF (scans until empty)',
 'Khay nạp giấy đang trống - đặt giấy vào rồi quét lại.': 'The document feeder is empty - load paper and scan again.',
 'Khi báo lỗi, nên gửi kèm ảnh chụp màn hình và file PDF gặp lỗi (nếu không chứa thông tin mật).': 'When reporting a '
                                                                                                   'bug, please '
                                                                                                   'attach a '
                                                                                                   'screenshot and '
                                                                                                   'the PDF file (if '
                                                                                                   'it contains no '
                                                                                                   'confidential '
                                                                                                   'data).',
 'Khi báo lỗi, nên gửi kèm ảnh chụp màn hình và file PDF gặp lỗi (nếu không chứa thông tin riêng tư, nhạy cảm).': 'When '
                                                                                                                  'reporting '
                                                                                                                  'a '
                                                                                                                  'bug, '
                                                                                                                  'please '
                                                                                                                  'attach '
                                                                                                                  'a '
                                                                                                                  'screenshot '
                                                                                                                  'and '
                                                                                                                  'the '
                                                                                                                  'PDF '
                                                                                                                  'file '
                                                                                                                  'that '
                                                                                                                  'caused '
                                                                                                                  'it '
                                                                                                                  '(if '
                                                                                                                  'it '
                                                                                                                  'contains '
                                                                                                                  'no '
                                                                                                                  'private '
                                                                                                                  'or '
                                                                                                                  'sensitive '
                                                                                                                  'information).',
 'Khi bấm Ký, driver của USB Token sẽ hỏi mã PIN. File đã ký được lưu thành file mới.\nSau khi ký, mọi chỉnh sửa và lưu đè sẽ làm chữ ký mất hiệu lực.': 'When '
                                                                                                                                                         'you '
                                                                                                                                                         'click '
                                                                                                                                                         'Sign, '
                                                                                                                                                         'the '
                                                                                                                                                         'USB '
                                                                                                                                                         'Token '
                                                                                                                                                         'driver '
                                                                                                                                                         'will '
                                                                                                                                                         'prompt '
                                                                                                                                                         'for '
                                                                                                                                                         'a '
                                                                                                                                                         'PIN. '
                                                                                                                                                         'The '
                                                                                                                                                         'signed '
                                                                                                                                                         'file '
                                                                                                                                                         'is '
                                                                                                                                                         'saved '
                                                                                                                                                         'as '
                                                                                                                                                         'a '
                                                                                                                                                         'new '
                                                                                                                                                         'file.\n'
                                                                                                                                                         'After '
                                                                                                                                                         'signing, '
                                                                                                                                                         'any '
                                                                                                                                                         'editing '
                                                                                                                                                         'or '
                                                                                                                                                         'overwriting '
                                                                                                                                                         'will '
                                                                                                                                                         'invalidate '
                                                                                                                                                         'the '
                                                                                                                                                         'signature.',
 'Khi căn sát mép: chừa lề 1 cm (bỏ tích để căn sát mép trang)': 'When aligning to an edge, keep a 1 cm margin '
                                                                 '(untick to align to the page edge)',
 'Khoanh đám mây': 'Cloud',
 'Khoảng cách giữa các trang': 'Gap between pages',
 'Khung chữ': 'Text box',
 'Khung chữ nhật': 'Rectangle frame',
 'Khung xem trước chỉ thêm khi máy chưa có (vd Microsoft Edge đã có sẵn thì giữ của Edge)': 'The preview pane is '
                                                                                            'only added when the PC '
                                                                                            'has none (e.g. '
                                                                                            "Microsoft Edge's "
                                                                                            'existing previewer is '
                                                                                            'kept)',
 'Khác': 'Other',
 'Không bắt buộc - dùng để gỡ giới hạn sau này. Để trống = tự sinh ngẫu nhiên': 'Optional - used to lift the '
                                                                                'restrictions later. Leave blank = '
                                                                                'random',
 'Không có bảng thì lấy toàn bộ chữ (mỗi dòng 1 hàng)': 'If there is no table, take all the text (one row per line)',
 'Không có chữ trong vùng - đã tô sáng cả vùng': 'No text in the area - the whole area has been highlighted',
 'Không có chữ trong vùng - đã đánh dấu cả vùng': 'No text in the area - marked the whole area',
 'Không có file phù hợp trong thư mục này': 'No matching files in this folder',
 'Không có ghi chú / nét vẽ tại đây': 'No comment / drawing here',
 'Không có thay đổi cần lưu': 'No changes to save',
 'Không có trang nào trùng nhau.': 'No duplicate pages found.',
 'Không có trang nào để in trong phạm vi đã chọn.': 'There are no pages to print in the selected range.',
 'Không cần mật khẩu: người nhận vẫn mở xem bình thường nhưng không in, không sao chép chữ / ảnh, không sửa được (tuỳ chọn).': 'No '
                                                                                                                               'password '
                                                                                                                               'needed: '
                                                                                                                               'recipients '
                                                                                                                               'can '
                                                                                                                               'still '
                                                                                                                               'open '
                                                                                                                               'and '
                                                                                                                               'read '
                                                                                                                               'the '
                                                                                                                               'file '
                                                                                                                               'but '
                                                                                                                               'cannot '
                                                                                                                               'print, '
                                                                                                                               'copy '
                                                                                                                               'text '
                                                                                                                               '/ '
                                                                                                                               'images '
                                                                                                                               'or '
                                                                                                                               'edit '
                                                                                                                               'it '
                                                                                                                               '(your '
                                                                                                                               'choice).',
 'Không cần nếu file chỉ bị giới hạn quyền': 'Not needed if the file is only permission-restricted',
 'Không di chuyển được ảnh này (ảnh nằm trong nhóm đối tượng của file gốc).': 'This image cannot be moved (it is '
                                                                              'inside an object group of the '
                                                                              'original file).',
 "Không ghi được file — file đang được mở bởi chương trình khác (Acrobat, trình duyệt...).\nHãy đóng file đó hoặc dùng 'Lưu thành'.": 'Could '
                                                                                                                                      'not '
                                                                                                                                      'write '
                                                                                                                                      'the '
                                                                                                                                      'file '
                                                                                                                                      '— '
                                                                                                                                      'it '
                                                                                                                                      'is '
                                                                                                                                      'open '
                                                                                                                                      'in '
                                                                                                                                      'another '
                                                                                                                                      'program '
                                                                                                                                      '(Acrobat, '
                                                                                                                                      'browser...).\n'
                                                                                                                                      'Please '
                                                                                                                                      'close '
                                                                                                                                      'that '
                                                                                                                                      'file '
                                                                                                                                      'or '
                                                                                                                                      'use '
                                                                                                                                      "'Save "
                                                                                                                                      "As'.",
 'Không ghi được file — file đang được mở bởi chương trình khác.': 'Could not write the file — it is open in another '
                                                                   'program.',
 'Không lưu được ảnh:': 'Could not save the image:',
 'Không mở được máy in.': 'Cannot open the printer.',
 'Không mở được thư mục:': 'Cannot open folder:',
 'Không nhận dạng được chữ nào trong vùng này.': 'No text was recognized in this region.',
 'Không nền': 'No fill',
 'Không theo bộ  (1,1,1  2,2,2  3,3,3)': 'Uncollated  (1,1,1  2,2,2  3,3,3)',
 'Không thấy chỗ trống nào để tạo ô form (dòng gạch dưới / chấm chấm, đường kẻ trống, ô vuông, khung trống).\n\nFile scan cần OCR trước, hoặc dùng nút Tạo ô form để kéo ô bằng tay.': 'No '
                                                                                                                                                                                       'blanks '
                                                                                                                                                                                       'found '
                                                                                                                                                                                       'to '
                                                                                                                                                                                       'turn '
                                                                                                                                                                                       'into '
                                                                                                                                                                                       'form '
                                                                                                                                                                                       'fields '
                                                                                                                                                                                       '(underscored '
                                                                                                                                                                                       '/ '
                                                                                                                                                                                       'dotted '
                                                                                                                                                                                       'lines, '
                                                                                                                                                                                       'empty '
                                                                                                                                                                                       'rules, '
                                                                                                                                                                                       'squares, '
                                                                                                                                                                                       'empty '
                                                                                                                                                                                       'boxes).\n'
                                                                                                                                                                                       '\n'
                                                                                                                                                                                       'Scanned '
                                                                                                                                                                                       'files '
                                                                                                                                                                                       'need '
                                                                                                                                                                                       'OCR '
                                                                                                                                                                                       'first, '
                                                                                                                                                                                       'or '
                                                                                                                                                                                       'use '
                                                                                                                                                                                       'Form '
                                                                                                                                                                                       'field '
                                                                                                                                                                                       'to '
                                                                                                                                                                                       'draw '
                                                                                                                                                                                       'fields '
                                                                                                                                                                                       'by '
                                                                                                                                                                                       'hand.',
 'Không thể xoá tất cả các trang.': 'Cannot delete all pages.',
 'Không tìm thấy chữ cần xoá.\n\nFile scan (chữ là ảnh) cần OCR trước, hoặc kéo vùng bằng nút Đánh dấu xoá.': 'No '
                                                                                                              'matching '
                                                                                                              'text '
                                                                                                              'found.\n'
                                                                                                              '\n'
                                                                                                              'Scanned '
                                                                                                              'files '
                                                                                                              '(text '
                                                                                                              'as '
                                                                                                              'images) '
                                                                                                              'need '
                                                                                                              'OCR '
                                                                                                              'first, '
                                                                                                              'or '
                                                                                                              'drag '
                                                                                                              'areas '
                                                                                                              'with '
                                                                                                              'Mark '
                                                                                                              'for '
                                                                                                              'redaction.',
 'Không tìm thấy con dấu theo màu mực đã chọn - thử chế độ tách nền khác.': 'No stamp found for the chosen ink color '
                                                                            '- try another background removal mode.',
 'Không tìm thấy công cụ phù hợp.': 'No matching tool found.',
 'Không tìm thấy dữ liệu giấy phép (thư mục licenses).': 'License data not found (licenses folder).',
 'Không tìm thấy file:': 'File not found:',
 'Không tìm thấy khoá bí mật (token chưa cắm hoặc sai token).': 'Private key not found (token not plugged in or '
                                                                'wrong token).',
 'Không tìm thấy khoá bí mật của chứng thư.': "Could not find the certificate's private key.",
 'Không tìm thấy khối này trên trang (có thể đã bị xoá).': 'Object not found on the page (it may have been deleted).',
 'Không tìm thấy máy scan nào (WIA). Kiểm tra máy đã cắm / bật và đã cài driver.': 'No scanner found (WIA). Check '
                                                                                   'that it is plugged in, switched '
                                                                                   'on and has a driver installed.',
 'Không tìm thấy thư mục:': 'Folder not found:',
 'Không tìm thấy thẻ / token trong đầu đọc.': 'No card / token found in the reader.',
 'Không tìm thấy token — hãy cắm USB Token.': 'Token not found — please plug in the USB Token.',
 'Không tìm thấy ô form này trên trang (có thể đã bị xoá).': 'This form field is not on the page (it may have been '
                                                             'deleted).',
 'Không tìm thấy địa chỉ web / email nào chưa có liên kết.': 'No unlinked web / email addresses found.',
 'Không tìm thấy ảnh này trong nội dung trang.': 'This image was not found in the page content.',
 'Không tìm thấy ảnh này trên trang.': 'This image is not on the page.',
 'Không tô nền': 'No background fill',
 'Không xem trước được loại file này': 'Cannot preview this file type',
 'Không xem trước được:\n': 'Cannot preview:\n',
 'Không xác định được lần vẽ ảnh này trong nội dung trang.': 'Could not locate where this image is drawn in the page '
                                                             'content.',
 'Không xử lý được file nào.': 'Could not process any files.',
 'Không đọc được ảnh nào.': 'Could not read any images.',
 'Không đọc được ảnh.': 'Cannot read the image.',
 'Không đọc được ảnh:': 'Could not read the image:',
 'Khổ': 'Size',
 'Khổ giấy': 'Paper size',
 'Khổ giấy và bố cục': 'Paper and layout',
 'Khổ trang': 'Page size',
 'Khổ trang (sách)': 'Page size (book)',
 'Khổ tờ in': 'Sheet size',
 'Kiểm tra': 'Check',
 'Kiểm tra chứng thư số': 'Certificate validation',
 'Kiểm tra chữ ký số': 'Verify digital signature',
 'Kiểm tra chữ ký trong file...': 'Verify signatures in file...',
 'Kiểm tra các chữ ký số trong file: người ký, thời điểm, file có bị sửa sau khi ký không.': 'Verify the digital '
                                                                                             'signatures in the '
                                                                                             'file: signer, time, '
                                                                                             'and whether the file '
                                                                                             'was modified after '
                                                                                             'signing.',
 'Kiểm tra lại các ô; ô thừa: chế độ Tạo ô form, chuột phải để xoá': 'Check the fields; to remove an extra one, '
                                                                     'right-click it in Form field mode',
 'Kiểu': 'Type',
 'Kiểu bút vẽ': 'Pen style',
 'Kiểu slide': 'Slide type',
 'Kiểu:': 'Style:',
 'KÝ SỐ: kéo vùng đặt chữ ký (Esc để huỷ)': 'SIGN: drag to place the signature area (Esc to cancel)',
 'Kéo chuột để vẽ / ký tự do': 'Drag to draw / sign freehand',
 'Kéo hoặc dùng phím mũi tên để di chuyển (Shift: bước 10 pt)': 'Drag or use the arrow keys to move (Shift: 10 pt '
                                                                'steps)',
 'Kéo hoặc dùng phím mũi tên để di chuyển, kéo ô vuông ở góc dưới phải để đổi kích thước': 'Drag or use the arrow '
                                                                                           'keys to move, drag the '
                                                                                           'square at the '
                                                                                           'bottom-right corner to '
                                                                                           'resize',
 'Kéo khung (hoặc bấm) trên trang để thêm chữ có định dạng': 'Drag a box (or click) on the page to add formatted '
                                                             'text',
 'Kéo khung (hoặc bấm) để thêm hộp chữ có viền / nền': 'Drag a frame (or click) to add a text box with border / fill',
 'Kéo qua chữ để chọn - Ctrl+C sao chép, chuột phải để tô sáng / gạch chân...': 'Drag over text to select - Ctrl+C '
                                                                                'to copy, right-click to highlight / '
                                                                                'underline...',
 'Kéo qua đoạn chữ cần gạch chân': 'Drag over the text to underline',
 'Kéo qua đoạn chữ cần gạch ngang (đề xuất xoá)': 'Drag over the text to strike out (suggest deletion)',
 'Kéo qua đoạn chữ cần tô sáng': 'Drag over the text to highlight',
 'Kéo thả file PDF vào đây\nhoặc bấm  Mở  (Ctrl+O)': 'Drag and drop a PDF file here\nor click  Open  (Ctrl+O)',
 'Kéo thả file vào đây\nhoặc bấm  Thêm file': 'Drag and drop files here\nor click  Add file',
 'Kéo thả thumbnail để đổi thứ tự, xoá, xoay, gộp trang': 'Drag and drop thumbnails to reorder, delete, rotate, or '
                                                          'merge pages',
 'Kéo từ điểm cần chỉ tới vị trí đặt hộp chữ': 'Drag from the point to indicate to where the text box goes',
 'Kéo vùng (hoặc bấm) để đặt chữ ký tay - vẽ bằng chuột hoặc dùng ảnh': 'Drag an area (or click) to place a '
                                                                        'handwritten signature - draw with the mouse '
                                                                        'or use an image',
 'Kéo vùng cần xoá vĩnh viễn (chữ, ảnh, hình bên dưới) - xem lại rồi bấm Xoá vĩnh viễn': 'Drag over an area to '
                                                                                         'redact permanently (text, '
                                                                                         'images, graphics '
                                                                                         'underneath) - review, then '
                                                                                         'click Redact permanently',
 'Kéo vùng trên trang để chụp thành ảnh (sao chép vào clipboard)': 'Drag an area of the page to capture it as an '
                                                                   'image (copied to the clipboard)',
 'Kéo vùng trên trang để đặt chữ ký số hiển thị': 'Drag an area on the page to place the visible digital signature',
 'Kéo vùng để tạo ô điền form (ô chữ, ô đánh dấu, danh sách chọn)': 'Drag an area to create a form field (text, '
                                                                    'checkbox, choice list)',
 'Kéo để di chuyển, kéo ô vuông ở góc dưới phải để đổi kích thước': 'Drag to move, drag the bottom-right square to '
                                                                    'resize',
 'Kéo để vẽ hình chữ nhật (hoặc che trắng)': 'Drag to draw a rectangle (or whiteout)',
 'Kéo để vẽ hình elip': 'Drag to draw an ellipse',
 'Kéo để vẽ khung đám mây (khoanh vùng cần sửa)': 'Drag to draw a cloud (mark an area to fix)',
 'Kéo để vẽ mũi tên': 'Drag to draw an arrow',
 'Kéo để vẽ đường thẳng': 'Drag to draw a line',
 'Kích thước': 'Size',
 'Kích thước thật (100%)': 'Actual size (100%)',
 'Kích thước trang': 'Page size',
 'Ký': 'Sign',
 'Ký bằng USB Token (VNPT-CA, Viettel-CA...), đặt vị trí chữ ký': 'Sign with a USB Token (VNPT-CA, Viettel-CA...), '
                                                                  'place the signature position',
 'Ký số': 'Digitally sign',
 'Ký số (kéo vùng đặt chữ ký)...': 'Digital signature (drag to place the signature area)...',
 'Ký số PDF': 'Sign PDF',
 'Ký số PDF (PAdES), dấu thời gian': 'PDF digital signatures (PAdES), timestamps',
 'Ký số bởi: ': 'Signed by: ',
 'Ký số không hiển thị...': 'Invisible digital signature...',
 'Ký vào đây bằng chuột / bút cảm ứng': 'Sign here with the mouse / stylus',
 'Kẹt giấy trong máy scan.': 'Paper jam in the scanner.',
 'Kẻ khung mỗi trang': 'Draw a border on each page',
 'Kẻ viền quanh mỗi trang (khi in nhiều trang / tờ)': 'Draw a border around each page (several pages per sheet)',
 'Kẻ vạch phân cách': 'Draw separator lines',
 'Kẻ đường phân cách': 'Draw a separator line',
 'Kết nối máy chủ dấu thời gian': 'Timestamp server connection',
 'Kết quả': 'Result',
 'Kết quả ký': 'Signing result',
 'Kết quả sẽ hiện ở đây': 'Results will appear here',
 'Kết quả áp dụng thẳng vào file đang sửa - có thể hoàn tác (Ctrl+Z).': 'The result is applied directly to the file '
                                                                        'being edited - can be undone (Ctrl+Z).',
 'LIÊN KẾT: kéo vùng trên chữ / ảnh có sẵn, hoặc bấm để thêm chữ liên kết mới': 'LINK: drag over existing '
                                                                                'text/image, or click to add new '
                                                                                'link text',
 'LIÊN KẾT: kéo vùng trên chữ / ảnh có sẵn, hoặc bấm để thêm chữ liên kết mới - kéo liên kết đã chèn để di chuyển, chuột phải để sửa / xoá': 'LINK: '
                                                                                                                                             'drag '
                                                                                                                                             'over '
                                                                                                                                             'existing '
                                                                                                                                             'text '
                                                                                                                                             '/ '
                                                                                                                                             'image, '
                                                                                                                                             'or '
                                                                                                                                             'click '
                                                                                                                                             'to '
                                                                                                                                             'add '
                                                                                                                                             'new '
                                                                                                                                             'link '
                                                                                                                                             'text '
                                                                                                                                             '- '
                                                                                                                                             'drag '
                                                                                                                                             'an '
                                                                                                                                             'inserted '
                                                                                                                                             'link '
                                                                                                                                             'to '
                                                                                                                                             'move '
                                                                                                                                             'it, '
                                                                                                                                             'right-click '
                                                                                                                                             'to '
                                                                                                                                             'edit '
                                                                                                                                             '/ '
                                                                                                                                             'delete',
 'Liên kết': 'Link',
 'Liên kết - kéo để di chuyển': 'Link - drag to move',
 'Liệt kê khổ giấy, kích thước (mm), hướng và góc xoay của từng trang; có thể xuất CSV.': 'List the paper size, '
                                                                                          'dimensions (mm), '
                                                                                          'orientation and rotation '
                                                                                          'of every page; can export '
                                                                                          'CSV.',
 'Loại': 'Type',
 'Loại file:': 'Files of type:',
 'Loại phần mềm': 'Software type',
 'Loại thông tin:': 'Information type:',
 'Loại ô:': 'Field type:',
 'Làm file trông như được scan từ giấy (xám, hơi nghiêng, có hạt).': 'Make the file look like it was scanned from '
                                                                     'paper (gray, slightly tilted, grainy).',
 'Làm lại': 'Redo',
 'Làm mới': 'Refresh',
 'Làm nét chữ': 'Sharpen text',
 'Làm nền (nằm dưới nội dung)': 'Set as background (behind content)',
 'Làm phẳng': 'Flatten',
 'Làm phẳng (form, ghi chú)': 'Flatten (forms, annotations)',
 'Làm phẳng sau khi điền (không sửa được nữa)': 'Flatten after filling (no longer editable)',
 'Làm sạch / xoá dữ liệu ẩn': 'Clean / remove hidden data',
 'Làm sạch dữ liệu ẩn': 'Clean hidden data',
 'Làm thẳng': 'Deskew',
 'Làm thẳng bản scan': 'Deskew scanned pages',
 'Làm thẳng trang scan': 'Deskew scanned pages',
 'Làm đậm, đồng màu mực (giống dấu mới đóng)': 'Strengthen and unify the ink (like a fresh stamp)',
 'Lên': 'Up',
 'Lên thư mục cha': 'Up one level',
 'Lý do ký:': 'Reason for signing:',
 'Lý do: ': 'Reason: ',
 'Lưu': 'Save',
 'Lưu chữ ký này để dùng lại lần sau (chỉ lưu trên máy này)': 'Save this signature for next time (stored on this '
                                                              'computer only)',
 'Lưu con dấu': 'Save stamp',
 'Lưu các file được nhúng trong PDF (và file đính kèm trong ghi chú) ra thư mục.': 'Save files embedded in the PDF '
                                                                                   '(and attachments in annotations) '
                                                                                   'to a folder.',
 'Lưu các trang đang chọn ra file mới': 'Save the selected pages to a new file',
 'Lưu cạnh file gốc': 'Save next to the original file',
 'Lưu danh sách ghi chú ra file HTML (mở bằng trình duyệt / Word) hoặc CSV (Excel)': 'Save the comment list as HTML '
                                                                                     '(browser / Word) or CSV '
                                                                                     '(Excel)',
 'Lưu file đã ký': 'Save signed file',
 'Lưu kết quả vào:': 'Save result to:',
 'Lưu mã nguồn PDF Toolkit': 'Save PDF Toolkit source code',
 'Lưu mã nguồn đầy đủ của PDF Toolkit (file zip) ra máy': 'Save the complete PDF Toolkit source code (zip file)',
 'Lưu mã nguồn...': 'Save source code...',
 'Lưu nguyên gốc các ảnh nhúng trong PDF ra thư mục.': 'Save the original embedded images from the PDF to a folder.',
 'Lưu thành': 'Save as',
 'Lưu thành ảnh': 'Save as image',
 'Lưu thành ảnh...': 'Save as image...',
 'Lưu vào thư mục:': 'Save to folder:',
 'Lưu vào thư viện con dấu': 'Save to stamp library',
 'Lưu vùng thành ảnh': 'Save region as image',
 'Lưu vùng thành ảnh...': 'Save region as image...',
 'Lưới trang': 'Page grid',
 'Lấy chữ trong PDF ra TXT, Markdown (dùng cho ChatGPT / AI), HTML hoặc JSON.': 'Extract text from the PDF to TXT, '
                                                                                'Markdown (for ChatGPT/AI), HTML, or '
                                                                                'JSON.',
 'Lấy các bảng trong PDF ra Excel (mỗi bảng 1 sheet) hoặc CSV.': 'Extract tables from the PDF to Excel (one sheet '
                                                                 'per table) or CSV.',
 'Lấy các trang chọn ra thành file mới.': 'Extract the selected pages into a new file.',
 'Lấy file đính kèm': 'Extract attachments',
 'Lấy thời điểm ký từ máy chủ tin cậy qua Internet - chữ ký vẫn hợp lệ sau khi chứng thư hết hạn': 'Get the signing '
                                                                                                   'time from a '
                                                                                                   'trusted '
                                                                                                   'timestamp server '
                                                                                                   'over the '
                                                                                                   'Internet - the '
                                                                                                   'signature '
                                                                                                   'remains valid '
                                                                                                   'after the '
                                                                                                   'certificate '
                                                                                                   'expires',
 'Lấy ảnh gốc': 'Extract images',
 'Lấy ảnh gốc trong PDF': 'Extract original images',
 'Lặp kín trang': 'Tile the entire page',
 'Lề': 'Margin',
 'Lề dưới': 'Bottom margin',
 'Lề phải': 'Right margin',
 'Lề trái': 'Left margin',
 'Lề trên': 'Top margin',
 'Lọc theo tên hoặc giấy phép...': 'Filter by name or license...',
 'Lồng mẫu nền': 'Overlay template',
 'Lồng mẫu nền / tiêu đề thư (letterhead)': 'Overlay letterhead / background template',
 'LỖI': 'ERROR',
 'Lỗi: ': 'Error: ',
 'Lớp': 'Layer',
 'Lớp nội dung': 'Content layers',
 'Lớp nội dung (layers)': 'Content layers',
 'Markdown -> PDF': 'Markdown -> PDF',
 'Menu chuột phải "PDF Toolkit" cho file PDF (nén, tách, chuyển Word / Excel, OCR, ký số...)': 'Right-click "PDF '
                                                                                               'Toolkit" menu for '
                                                                                               'PDF files (compress, '
                                                                                               'split, convert to '
                                                                                               'Word / Excel, OCR, '
                                                                                               'sign...)',
 'MÁY CHỮ: bấm vào chỗ cần gõ chữ': 'TYPEWRITER: click where to type',
 'Màu': 'Color',
 'Màu && nét bút...': 'Pen color && width...',
 'Màu && độ dày nét...': 'Color && line width...',
 'Màu chữ': 'Text color',
 'Màu chữ mới': 'New text color',
 'Màu khác...': 'Other color...',
 'Màu mực:': 'Ink color:',
 'Màu nền': 'Background color',
 'Màu nền trang': 'Page background color',
 'Màu sắc': 'Color',
 'Màu trắng': 'White',
 'Màu viền:': 'Border color:',
 'Màu đen': 'Black',
 'Màu:': 'Color:',
 'Máy chủ dấu thời gian': 'Timestamp server',
 'Máy chữ': 'Typewriter',
 'Máy in': 'Printer',
 'Máy in này không hỗ trợ in 2 mặt tự động': 'This printer does not support automatic 2-sided printing',
 'Máy scan này không có khay nạp giấy tự động (ADF) - chọn Mặt kính.': 'This scanner has no automatic document '
                                                                       'feeder (ADF) - choose Flatbed.',
 'Máy scan đang bận - thử lại sau.': 'The scanner is busy - try again later.',
 'Máy scan đang tắt hoặc chưa kết nối.': 'The scanner is off or not connected.',
 'Máy scan → PDF': 'Scanner → PDF',
 'Máy tính chưa cài máy in nào.': 'No printer is installed on this computer.',
 'Mã hoá': 'Cryptography',
 'Mô tả vấn đề:': 'Describe the problem:',
 'MŨI TÊN: kéo từ gốc đến đầu mũi tên': 'ARROW: drag from the start to the arrowhead',
 'MŨI TÊN: kéo từ gốc đến đầu mũi tên - kéo viền khung mũi tên đã vẽ để di chuyển': 'ARROW: drag from tail to head - '
                                                                                    'drag the frame edge of a drawn '
                                                                                    'arrow to move it',
 'Mũi tên': 'Arrow',
 'Mạnh': 'Strong',
 'Mạnh - ảnh 96 dpi': 'Strong - 96 dpi images',
 'Mẫu tuỳ chỉnh': 'Custom template',
 'MẬT': 'CONFIDENTIAL',
 'Mật khẩu': 'Password',
 'Mật khẩu chủ (đổi quyền)': 'Owner password (change permissions)',
 'Mật khẩu chủ được tạo ngẫu nhiên (không lưu lại).': 'The owner password is randomly generated (not saved).',
 'Mật khẩu hiện tại': 'Current password',
 'Mật khẩu mở file': 'File-open password',
 'Mật khẩu mở file (nếu có)': 'File-open password (if any)',
 'Mặt kính (từng trang, hỏi quét tiếp)': 'Flatbed (page by page, asks to continue)',
 'Mọi số CCCD / CMND (9 hoặc 12 số)': 'All ID card numbers (9 or 12 digits)',
 'Mọi số tài khoản / dãy số dài (10-19 số)': 'All account numbers / long digit strings (10-19 digits)',
 'Mọi số tài khoản / dãy số dài (10-19 số, có thể cách nhau)': 'All account numbers / long digit strings (10-19 '
                                                               'digits, may be spaced)',
 'Mọi số điện thoại': 'All phone numbers',
 'Mọi trang': 'All pages',
 'Mọi trang trong phạm vi': 'All pages in range',
 'Mọi địa chỉ email': 'All email addresses',
 'Mỗi N trang một file': 'One file per N pages',
 'Mỗi dòng 1 cụm chữ, vd:\nNguyễn Văn A\nCông ty ABC': 'One phrase per line, e.g.:\nJohn Smith\nABC Company',
 'Mỗi dòng một lựa chọn': 'One choice per line',
 'Mỗi dòng: tên file | mô tả. Xoá dòng nào thì file đó bị xoá khỏi PDF.': 'One line per file: file name | '
                                                                          'description. Delete a line to remove that '
                                                                          'file from the PDF.',
 'Mỗi dòng: tên lớp | hiện hoặc ẩn': 'One per line: layer name | show or hide',
 'Mỗi dòng: từ trang | kiểu | tiền tố | số bắt đầu. Kiểu: 1 (1,2,3), i (i,ii), I (I,II), a (a,b), A (A,B), - (chỉ tiền tố)': 'One '
                                                                                                                             'rule '
                                                                                                                             'per '
                                                                                                                             'line: '
                                                                                                                             'from '
                                                                                                                             'page '
                                                                                                                             '| '
                                                                                                                             'style '
                                                                                                                             '| '
                                                                                                                             'prefix '
                                                                                                                             '| '
                                                                                                                             'start '
                                                                                                                             'number. '
                                                                                                                             'Style: '
                                                                                                                             '1 '
                                                                                                                             '(1,2,3), '
                                                                                                                             'i '
                                                                                                                             '(i,ii), '
                                                                                                                             'I '
                                                                                                                             '(I,II), '
                                                                                                                             'a '
                                                                                                                             '(a,b), '
                                                                                                                             'A '
                                                                                                                             '(A,B), '
                                                                                                                             '- '
                                                                                                                             '(prefix '
                                                                                                                             'only)',
 'Mỗi trang là 1 ảnh (giống bản gốc 100%, không sửa chữ)': 'Each page as one image (100% faithful, text not '
                                                           'editable)',
 'Mỗi trang một file': 'One file per page',
 'Một lần giữa trang': 'Once, in the middle of the page',
 'Mở': 'Open',
 'Mở PDF kết quả trong tab mới': 'Open the result PDF in a new tab',
 'Mở bảng trang (F4)': 'Show page panel (F4)',
 'Mở bằng PDF Toolkit': 'Open with PDF Toolkit',
 'Mở file PDF': 'Open PDF file',
 'Mở file kết quả': 'Open result file',
 'Mở file để xem, sửa chữ, chèn, sắp xếp trang, OCR, ký số...': 'Open a file to view, edit text, insert, arrange '
                                                                'pages, OCR, sign...',
 'Mở hộp thoại in của Windows (khay giấy, chất lượng in... riêng của máy in) - bấm In ở đó sẽ in ngay': 'Open the '
                                                                                                        'Windows '
                                                                                                        'print '
                                                                                                        'dialog '
                                                                                                        '(paper '
                                                                                                        'tray, print '
                                                                                                        'quality... '
                                                                                                        'specific to '
                                                                                                        'the '
                                                                                                        'printer) - '
                                                                                                        'clicking '
                                                                                                        'Print there '
                                                                                                        'prints '
                                                                                                        'right away',
 'Mở hộp thoại in của Windows - bấm In ở đó sẽ in ngay': 'Opens the Windows print dialog - clicking Print there '
                                                         'prints right away',
 'Mở khoá': 'Unlock',
 'Mở khoá / gỡ giới hạn': 'Unlock / remove restrictions',
 'Mở kết quả trong tab mới': 'Open result in a new tab',
 'Mở liên kết': 'Open link',
 'Mở thêm file PDF trong tab mới  (Ctrl+O)': 'Open more PDF files in a new tab  (Ctrl+O)',
 'Mở thư mới (Outlook / ứng dụng email mặc định) đính kèm file PDF đang mở': 'Open a new email (Outlook / default '
                                                                             'mail app) with the current PDF '
                                                                             'attached',
 'Mở thư mục': 'Open folder',
 'Mở thư mục chứa file': 'Open containing folder',
 'Mở thư mục giấy phép': 'Open licenses folder',
 'Mở thư mục kết quả': 'Open result folder',
 'Mở thư mục thư viện': 'Open library folder',
 'Mở trong tab mới': 'Open in a new tab',
 'Mở trong trình sửa': 'Open in editor',
 'MỤC LỤC': 'TABLE OF CONTENTS',
 'Mức nén': 'Compression level',
 'Mực tím': 'Purple ink',
 'Mực xanh': 'Blue ink',
 'Mực đen / xám (bỏ nền giấy)': 'Black / gray ink (drop the paper)',
 'Mực đỏ': 'Red ink',
 'Ngang': 'Horizontal',
 'Ngang trước (trái→phải)': 'Horizontal first (left→right)',
 'Nguồn giấy:': 'Paper source:',
 'Nguồn ảnh': 'Source',
 'Ngày': 'Date',
 'Ngày ký: ': 'Signing date: ',
 'Ngày sửa': 'Date modified',
 'Ngôn ngữ': 'Language',
 'Ngôn ngữ giao diện': 'Interface language',
 'Người ký': 'Signer',
 'Người ký, thời điểm, file có bị sửa sau khi ký không': 'Signer, time, and whether the file was modified after '
                                                         'signing',
 'Người nhận vẫn mở xem được bình thường; Adobe Reader, Foxit, Edge, Chrome... sẽ chặn các quyền đã bỏ chọn.': 'Recipients '
                                                                                                               'can '
                                                                                                               'still '
                                                                                                               'open '
                                                                                                               'the '
                                                                                                               'file '
                                                                                                               'normally; '
                                                                                                               'Adobe '
                                                                                                               'Reader, '
                                                                                                               'Foxit, '
                                                                                                               'Edge, '
                                                                                                               'Chrome... '
                                                                                                               'block '
                                                                                                               'the '
                                                                                                               'unticked '
                                                                                                               'permissions.',
 'Người tạo': 'Author',
 'Nhiều trang / 1 tờ': 'Pages per sheet',
 'Nhiều trang / tờ': 'Multiple pages per sheet',
 'Nhân bản': 'Duplicate',
 'Nhân bản trang': 'Duplicate pages',
 'Nhãn trang': 'Page labels',
 'Nhãn trang (i, ii, A-1...)': 'Page labels (i, ii, A-1...)',
 'Nhúng file bất kỳ (Excel, Word, ảnh...) vào bên trong PDF.': 'Embed any file (Excel, Word, images...) inside the '
                                                               'PDF.',
 'Nhúng file đính kèm vào PDF': 'Embed attachment into PDF',
 'Nhận biết chữ đậm': 'Detect bold text',
 'Nhận dạng bản scan rồi dựng lại trang bằng CHỮ THẬT (giữ vị trí, cỡ, màu, chữ đậm) - sửa chữ trực tiếp trong trình sửa. Logo, con dấu, chữ ký, kẻ bảng được giữ lại trên nền.': 'Recognizes '
                                                                                                                                                                                  'a '
                                                                                                                                                                                  'scan '
                                                                                                                                                                                  'and '
                                                                                                                                                                                  'rebuilds '
                                                                                                                                                                                  'each '
                                                                                                                                                                                  'page '
                                                                                                                                                                                  'with '
                                                                                                                                                                                  'REAL '
                                                                                                                                                                                  'text '
                                                                                                                                                                                  '(keeps '
                                                                                                                                                                                  'position, '
                                                                                                                                                                                  'size, '
                                                                                                                                                                                  'color, '
                                                                                                                                                                                  'bold) '
                                                                                                                                                                                  '- '
                                                                                                                                                                                  'edit '
                                                                                                                                                                                  'the '
                                                                                                                                                                                  'text '
                                                                                                                                                                                  'directly '
                                                                                                                                                                                  'in '
                                                                                                                                                                                  'the '
                                                                                                                                                                                  'editor. '
                                                                                                                                                                                  'Logos, '
                                                                                                                                                                                  'stamps, '
                                                                                                                                                                                  'signatures '
                                                                                                                                                                                  'and '
                                                                                                                                                                                  'table '
                                                                                                                                                                                  'lines '
                                                                                                                                                                                  'stay '
                                                                                                                                                                                  'on '
                                                                                                                                                                                  'the '
                                                                                                                                                                                  'background.',
 'Nhận dạng chữ': 'Text recognition',
 'Nhận dạng chữ (OCR) - nằm trong MuPDF': 'Text recognition (OCR) - part of MuPDF',
 'Nhận dạng chữ (OCR) trong vùng...': 'Recognize text (OCR) in region...',
 'Nhận ra mực con dấu màu: ': 'Detected stamp ink color: ',
 'Nhận xét': 'Comment',
 'Nhập chữ cần xoá hoặc chọn ít nhất một loại thông tin.': 'Enter text to redact or pick at least one information '
                                                           'type.',
 'Nhập nội dung...': 'Enter text...',
 'Nhập nội dung... (Enter để xuống dòng, bôi đen để định dạng)': 'Enter content... (Enter for a new line, select '
                                                                 'text to format)',
 'Nhẹ': 'Light',
 'Nhẹ - giữ nguyên chất lượng ảnh': 'Light - keeps image quality',
 'Nén': 'Compress',
 'Nén PDF': 'Compress PDF',
 'Nét vẽ tay': 'Freehand',
 'Nét:': 'Stroke:',
 'Nơi ký:': 'Location:',
 'Nơi ký: ': 'Location: ',
 'Nằm dưới nội dung': 'Behind content',
 'Nền trang': 'Page background',
 'Nền trắng - chỉ giữ chữ (file nhẹ nhất)': 'White background - text only (smallest file)',
 'Nền trắng, không viền - dùng để che nội dung bên dưới': 'White background, no border - used to cover content '
                                                          'underneath',
 'Nền:': 'Fill:',
 'Nối tất cả các trang thành một trang dài liền mạch (dọc hoặc ngang) - tiện đọc trên màn hình.': 'Join all pages '
                                                                                                  'into one '
                                                                                                  'continuous long '
                                                                                                  'page (vertical or '
                                                                                                  'horizontal) - '
                                                                                                  'handy for reading '
                                                                                                  'on screen.',
 'Nội dung': 'Content',
 'Nội dung cần chèn:': 'Text to insert:',
 'Nội dung dấu': 'Stamp content',
 'Nội dung ghi chú cho đoạn chữ:': 'Note for the selected text:',
 'Nội dung ghi chú:': 'Note content:',
 'Nội dung trang - ảnh nằm trong trang như ảnh gốc của file': 'Page content - the image becomes part of the page '
                                                              "like the file's own images",
 'Nội dung đề xuất chèn vào:': 'Text to suggest inserting:',
 'Nội dung:': 'Content:',
 'Nới rộng trang thêm lề trắng (vd để đóng gáy, ghi chú bên lề) - nội dung, ghi chú, liên kết giữ nguyên.': 'Enlarge '
                                                                                                            'pages '
                                                                                                            'with '
                                                                                                            'white '
                                                                                                            'margins '
                                                                                                            '(e.g. '
                                                                                                            'for '
                                                                                                            'binding, '
                                                                                                            'side '
                                                                                                            'notes) '
                                                                                                            '- '
                                                                                                            'content, '
                                                                                                            'comments '
                                                                                                            'and '
                                                                                                            'links '
                                                                                                            'stay '
                                                                                                            'unchanged.',
 'Nửa sau có thứ tự ngược (lật cả xấp giấy để scan)': 'Second half is in reverse order (whole stack flipped for '
                                                      'scanning)',
 'OCR && Tối ưu': 'OCR && Optimize',
 'OCR - nhận dạng chữ': 'OCR - recognize text',
 'OCR - nhận dạng chữ bản scan': 'OCR - recognize text in scans',
 'OCR bản scan': 'OCR scanned document',
 'OCR lại tất cả các trang': 'Re-run OCR on all pages',
 'OCR sửa được': 'Editable OCR',
 'OCR → PDF chỉnh sửa được': 'OCR → editable PDF',
 'OCR → PDF sửa được': 'OCR → editable PDF',
 'Office add-in (Word, Excel, PowerPoint)': 'Office add-in (Word, Excel, PowerPoint)',
 'Office không xuất được PDF': 'Office failed to export PDF',
 'PDF -> Excel': 'PDF -> Excel',
 'PDF -> Word': 'PDF -> Word',
 'PDF / Ảnh (*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp);;PDF (*.pdf);;Tất cả (*.*)': 'PDF / Image (*.pdf '
                                                                                              '*.png *.jpg *.jpeg '
                                                                                              '*.bmp *.tif *.tiff '
                                                                                              '*.webp);;PDF '
                                                                                              '(*.pdf);;All files '
                                                                                              '(*.*)',
 'PDF Toolkit (chương trình này)': 'PDF Toolkit (this program)',
 'PDF Toolkit được xây dựng trên các thành phần mã nguồn mở dưới đây, mỗi thành phần thuộc bản quyền của tác giả tương ứng và được dùng theo giấy phép ghi kèm.': 'PDF '
                                                                                                                                                                  'Toolkit '
                                                                                                                                                                  'is '
                                                                                                                                                                  'built '
                                                                                                                                                                  'on '
                                                                                                                                                                  'the '
                                                                                                                                                                  'open-source '
                                                                                                                                                                  'components '
                                                                                                                                                                  'below. '
                                                                                                                                                                  'Each '
                                                                                                                                                                  'component '
                                                                                                                                                                  'is '
                                                                                                                                                                  'copyrighted '
                                                                                                                                                                  'by '
                                                                                                                                                                  'its '
                                                                                                                                                                  'respective '
                                                                                                                                                                  'authors '
                                                                                                                                                                  'and '
                                                                                                                                                                  'used '
                                                                                                                                                                  'under '
                                                                                                                                                                  'the '
                                                                                                                                                                  'license '
                                                                                                                                                                  'shown.',
 'PDF giới hạn khổ trang 5 m; nội dung dài hơn sẽ được thu nhỏ cho vừa.': 'PDF pages are limited to 5 m; longer '
                                                                          'content is scaled down to fit.',
 'PDF scan (chỉ có ảnh) sẽ ra Word chứa ảnh. File phức tạp có thể lệch bố cục.': 'Scanned PDFs (image-only) will '
                                                                                 'produce a Word file containing '
                                                                                 'images. Complex files may have '
                                                                                 'layout shifts.',
 'PDF scan không có lớp chữ sẽ ra file rỗng (cần OCR).': 'A scanned PDF without a text layer will produce an empty '
                                                         'file (OCR required).',
 'PDF scan: hãy OCR trước để có chữ sửa được. Font không có trên máy sẽ được PowerPoint thay bằng font tương tự.': 'Scanned '
                                                                                                                   'PDF: '
                                                                                                                   'run '
                                                                                                                   'OCR '
                                                                                                                   'first '
                                                                                                                   'to '
                                                                                                                   'get '
                                                                                                                   'editable '
                                                                                                                   'text. '
                                                                                                                   'Fonts '
                                                                                                                   'missing '
                                                                                                                   'on '
                                                                                                                   'the '
                                                                                                                   'computer '
                                                                                                                   'are '
                                                                                                                   'replaced '
                                                                                                                   'by '
                                                                                                                   'similar '
                                                                                                                   'ones '
                                                                                                                   'in '
                                                                                                                   'PowerPoint.',
 'PDF → PowerPoint': 'PDF → PowerPoint',
 'PDF → Văn bản / Markdown': 'PDF → Text / Markdown',
 'PDF → Ảnh': 'PDF → Image',
 'Phiên bản': 'Version',
 'Phân biệt chữ hoa / thường': 'Case sensitive',
 'Phê duyệt': 'Approval',
 'Phóng to': 'Zoom in',
 'Phóng to 1 trang ra nhiều tờ (vd 2x2 tờ A4) để in ghép thành poster.': 'Enlarge one page across multiple sheets '
                                                                         '(e.g. 2x2 A4 sheets) to print and assemble '
                                                                         'as a poster.',
 'Phóng to chữ trong ô soạn cho dễ đọc (không đổi cỡ chữ thật trên PDF)': 'Enlarge text in the editing box for '
                                                                          'readability (does not change the actual '
                                                                          'font size on the PDF)',
 'Phóng to vùng': 'Zoom to region',
 'Phù hợp nhất với bản scan có chữ. Ghi chú / liên kết trên trang được xoay sẽ bị bỏ.': 'Works best on scans '
                                                                                        'containing text. Notes / '
                                                                                        'links on rotated pages are '
                                                                                        'removed.',
 'Phạm vi': 'Range',
 'Phạm vi:': 'Range:',
 'Phần mềm PDF': 'PDF software',
 'Phần mềm nội bộ, chạy offline. Sử dụng các thành phần mã nguồn mở - xem Giấy phép bên thứ ba.': 'Internal '
                                                                                                  'software, runs '
                                                                                                  'offline. Uses '
                                                                                                  'open-source '
                                                                                                  'components - see '
                                                                                                  'Third-party '
                                                                                                  'licenses.',
 'Phần mềm nội bộ, dùng trong công ty': 'Internal software, for use within the company',
 'Quay lại': 'Back to',
 'Quay lại file đang làm việc (Esc)': 'Return to the open file (Esc)',
 'Quy tắc': 'Rules',
 'Quét': 'Scan',
 'Quét thêm trang từ máy scan, chèn sau trang đang xem': 'Scan more pages and insert them after the current page',
 'Quét tài liệu giấy thành PDF: mặt kính hoặc khay nạp giấy tự động': 'Scan paper documents to PDF: flatbed or '
                                                                      'automatic document feeder',
 'Quét tài liệu giấy từ máy scan thành PDF (mở trong tab mới)': 'Scan paper documents to PDF (opens in a new tab)',
 'Quản lý': 'Manage',
 'Rộng (mm)': 'Width (mm)',
 'Sao chép': 'Copy',
 'Sao chép chữ': 'Copy text',
 'Sao chép chữ trong vùng': 'Copy text in region',
 'Sao chép đường dẫn': 'Copy path',
 'Sao chép địa chỉ': 'Copy address',
 'Sao chép địa chỉ email': 'Copy email address',
 'Sao chép ảnh vùng chọn\tCtrl+C': 'Copy region as image\tCtrl+C',
 'Sau mỗi trang': 'After each page',
 'Sau trang': 'After page',
 'Sau trang số': 'After page number',
 'Shortcut ngoài Desktop': 'Desktop shortcut',
 'Shortcut trong Start Menu': 'Start Menu shortcut',
 'So sánh': 'Compare',
 'So sánh 2 file': 'Compare 2 files',
 'So sánh 2 file PDF': 'Compare 2 PDF files',
 'Sách gập': 'Booklet',
 'Sáng': 'Light',
 'Sáng / tương phản': 'Brightness / contrast',
 'Sắp bản scan 2 mặt': 'Collate duplex scan',
 'Sắp trang để in 2 mặt, gập đôi thành cuốn sách nhỏ (2 trang A5 trên 1 tờ A4 ngang).': 'Arrange pages for '
                                                                                        'double-sided printing, '
                                                                                        'folded into a small booklet '
                                                                                        '(2 A5 pages on one '
                                                                                        'landscape A4 sheet).',
 'Sắp xếp': 'Arrange',
 'Sắp xếp & Quản lý trang': 'Arrange & Manage Pages',
 'Sắp xếp bản scan 2 mặt (1 file)': 'Collate duplex scan (1 file)',
 'Sắp xếp trang trực quan': 'Visual page arrangement',
 'Số Bates': 'Bates number',
 'Số CCCD / CMND (9 hoặc 12 số)': 'ID card number (9 or 12 digits)',
 'Số bản': 'Copies',
 'Số bản:': 'Copies:',
 'Số bắt đầu': 'Starting number',
 'Số chữ số': 'Number of digits',
 'Số cột': 'Number of columns',
 'Số dương = xoay ngược chiều kim đồng hồ': 'Positive = counter-clockwise',
 'Số hàng': 'Number of rows',
 'Số phần': 'Number of parts',
 'Số trang': 'Page number',
 'Số trang / tờ': 'Pages per sheet',
 'Số trang mỗi file': 'Pages per file',
 'Số trang trắng': 'Number of blank pages',
 "Số trang tính từ 1. 'le' = trang lẻ, 'chan' = trang chẵn, '8-' = từ trang 8 đến cuối": 'Pages are numbered from 1. '
                                                                                         "'le' = odd pages, 'chan' = "
                                                                                         "even pages, '8-' = from "
                                                                                         'page 8 to the end',
 'Số tài khoản / dãy số dài (10-19 số, có thể cách nhau)': 'Account number / long digit string (10-19 digits, may be '
                                                           'spaced)',
 'Số điện thoại': 'Phone number',
 'SỬA CHỮ: bấm vào dòng chữ có khung xanh': 'EDIT TEXT: click the text line with the blue box',
 'SỬA FILE: bấm dòng chữ (khung xanh) để sửa chữ - kéo ảnh / chữ ký / con dấu / khối vẽ để di chuyển, góc dưới phải để đổi cỡ - bấm chọn rồi dùng nhóm Căn chỉnh, Delete để xoá': 'EDIT '
                                                                                                                                                                                  'FILE: '
                                                                                                                                                                                  'click '
                                                                                                                                                                                  'a '
                                                                                                                                                                                  'text '
                                                                                                                                                                                  'line '
                                                                                                                                                                                  '(blue '
                                                                                                                                                                                  'box) '
                                                                                                                                                                                  'to '
                                                                                                                                                                                  'edit '
                                                                                                                                                                                  'text '
                                                                                                                                                                                  '- '
                                                                                                                                                                                  'drag '
                                                                                                                                                                                  'images '
                                                                                                                                                                                  '/ '
                                                                                                                                                                                  'signatures '
                                                                                                                                                                                  '/ '
                                                                                                                                                                                  'stamps '
                                                                                                                                                                                  '/ '
                                                                                                                                                                                  'shapes '
                                                                                                                                                                                  'to '
                                                                                                                                                                                  'move, '
                                                                                                                                                                                  'bottom-right '
                                                                                                                                                                                  'corner '
                                                                                                                                                                                  'to '
                                                                                                                                                                                  'resize '
                                                                                                                                                                                  '- '
                                                                                                                                                                                  'click '
                                                                                                                                                                                  'to '
                                                                                                                                                                                  'select '
                                                                                                                                                                                  'then '
                                                                                                                                                                                  'use '
                                                                                                                                                                                  'the '
                                                                                                                                                                                  'Align '
                                                                                                                                                                                  'group, '
                                                                                                                                                                                  'Delete '
                                                                                                                                                                                  'to '
                                                                                                                                                                                  'remove',
 'SỬA FILE: bấm để chọn khung chữ / ảnh / chữ ký / con dấu / khối vẽ - giữ chuột kéo hoặc phím mũi tên để di chuyển (Shift: 10 pt), bấm đúp / Enter để sửa chữ, Delete để xoá, nhóm Căn chỉnh để căn': 'EDIT '
                                                                                                                                                                                                       'FILE: '
                                                                                                                                                                                                       'click '
                                                                                                                                                                                                       'to '
                                                                                                                                                                                                       'select '
                                                                                                                                                                                                       'a '
                                                                                                                                                                                                       'text '
                                                                                                                                                                                                       'box '
                                                                                                                                                                                                       '/ '
                                                                                                                                                                                                       'image '
                                                                                                                                                                                                       '/ '
                                                                                                                                                                                                       'signature '
                                                                                                                                                                                                       '/ '
                                                                                                                                                                                                       'stamp '
                                                                                                                                                                                                       '/ '
                                                                                                                                                                                                       'shape '
                                                                                                                                                                                                       '- '
                                                                                                                                                                                                       'drag '
                                                                                                                                                                                                       'or '
                                                                                                                                                                                                       'use '
                                                                                                                                                                                                       'the '
                                                                                                                                                                                                       'arrow '
                                                                                                                                                                                                       'keys '
                                                                                                                                                                                                       'to '
                                                                                                                                                                                                       'move '
                                                                                                                                                                                                       '(Shift: '
                                                                                                                                                                                                       '10 '
                                                                                                                                                                                                       'pt), '
                                                                                                                                                                                                       'double-click '
                                                                                                                                                                                                       '/ '
                                                                                                                                                                                                       'Enter '
                                                                                                                                                                                                       'to '
                                                                                                                                                                                                       'edit '
                                                                                                                                                                                                       'text, '
                                                                                                                                                                                                       'Delete '
                                                                                                                                                                                                       'to '
                                                                                                                                                                                                       'remove, '
                                                                                                                                                                                                       'Align '
                                                                                                                                                                                                       'group '
                                                                                                                                                                                                       'to '
                                                                                                                                                                                                       'align',
 'Sửa / xoá file đính kèm': 'Edit / remove attachments',
 'Sửa / xoá đính kèm': 'Edit / remove attachments',
 'Sửa PDF': 'Edit PDF',
 'Sửa bookmark (mục lục điều hướng)': 'Edit bookmarks (navigation outline)',
 'Sửa chữ': 'Edit text',
 'Sửa chữ...  (bấm đúp / Enter)': 'Edit text...  (double-click / Enter)',
 'Sửa dòng chữ': 'Edit text line',
 'Sửa file': 'Edit file',
 'Sửa file PDF lỗi': 'Repair broken PDF file',
 'Sửa file lỗi': 'Repair broken file',
 'Sửa khối chữ': 'Edit text block',
 'Sửa liên kết': 'Edit link',
 'Sửa mọi thành phần trong file: bấm chọn khung chữ / ảnh / chữ ký / con dấu / khối vẽ, giữ chuột kéo hoặc phím mũi tên để di chuyển, bấm đúp để sửa chữ, căn chỉnh, Delete để xoá': 'Edit '
                                                                                                                                                                                     'every '
                                                                                                                                                                                     'element '
                                                                                                                                                                                     'in '
                                                                                                                                                                                     'the '
                                                                                                                                                                                     'file: '
                                                                                                                                                                                     'click '
                                                                                                                                                                                     'to '
                                                                                                                                                                                     'select '
                                                                                                                                                                                     'a '
                                                                                                                                                                                     'text '
                                                                                                                                                                                     'box '
                                                                                                                                                                                     '/ '
                                                                                                                                                                                     'image '
                                                                                                                                                                                     '/ '
                                                                                                                                                                                     'signature '
                                                                                                                                                                                     '/ '
                                                                                                                                                                                     'stamp '
                                                                                                                                                                                     '/ '
                                                                                                                                                                                     'shape, '
                                                                                                                                                                                     'drag '
                                                                                                                                                                                     'or '
                                                                                                                                                                                     'use '
                                                                                                                                                                                     'the '
                                                                                                                                                                                     'arrow '
                                                                                                                                                                                     'keys '
                                                                                                                                                                                     'to '
                                                                                                                                                                                     'move, '
                                                                                                                                                                                     'double-click '
                                                                                                                                                                                     'to '
                                                                                                                                                                                     'edit '
                                                                                                                                                                                     'text, '
                                                                                                                                                                                     'align, '
                                                                                                                                                                                     'Delete '
                                                                                                                                                                                     'to '
                                                                                                                                                                                     'remove',
 'Sửa mọi thành phần trong file: bấm dòng chữ để sửa chữ; kéo ảnh / chữ ký / con dấu / khối vẽ / liên kết để di chuyển, đổi cỡ; bấm chọn rồi căn chỉnh, Delete để xoá': 'Edit '
                                                                                                                                                                        'everything '
                                                                                                                                                                        'in '
                                                                                                                                                                        'the '
                                                                                                                                                                        'file: '
                                                                                                                                                                        'click '
                                                                                                                                                                        'a '
                                                                                                                                                                        'text '
                                                                                                                                                                        'line '
                                                                                                                                                                        'to '
                                                                                                                                                                        'edit '
                                                                                                                                                                        'text; '
                                                                                                                                                                        'drag '
                                                                                                                                                                        'images '
                                                                                                                                                                        '/ '
                                                                                                                                                                        'signatures '
                                                                                                                                                                        '/ '
                                                                                                                                                                        'stamps '
                                                                                                                                                                        '/ '
                                                                                                                                                                        'shapes '
                                                                                                                                                                        '/ '
                                                                                                                                                                        'links '
                                                                                                                                                                        'to '
                                                                                                                                                                        'move '
                                                                                                                                                                        'or '
                                                                                                                                                                        'resize; '
                                                                                                                                                                        'click '
                                                                                                                                                                        'to '
                                                                                                                                                                        'select '
                                                                                                                                                                        'then '
                                                                                                                                                                        'align, '
                                                                                                                                                                        'Delete '
                                                                                                                                                                        'to '
                                                                                                                                                                        'remove',
 'Sửa mục lục điều hướng (bookmark)': 'Edit bookmarks (navigation outline)',
 'Sửa địa chỉ liên kết...': 'Edit link address...',
 'THÊM CHỮ: kéo khung (hoặc bấm) tại vị trí cần thêm chữ': 'ADD TEXT: drag a box (or click) where you want to add '
                                                           'text',
 'Tab sau': 'Next tab',
 'Tab trước': 'Previous tab',
 'Tham số ký không hợp lệ.': 'Invalid signing parameters.',
 'Thay bằng:': 'Replace with:',
 'Thay bằng: ': 'Replace with: ',
 'Thay các trang đang chọn bằng trang lấy từ file PDF khác': 'Replace the selected pages with pages from another PDF',
 'Thay trang...': 'Replace pages...',
 'Thay tất cả': 'Replace all',
 'Theo Windows': 'Follow Windows',
 'Theo bookmark cấp 1': 'By level-1 bookmarks',
 'Theo bộ  (1,2,3  1,2,3  1,2,3)': 'Collated  (1,2,3  1,2,3  1,2,3)',
 'Theo dung lượng mong muốn': 'By target file size',
 'Theo khoảng trang': 'By page range',
 'Theo kích thước ảnh': 'By image size',
 'Theo số mm': 'By millimeters',
 'Thiếu thành phần pywin32 để dùng máy scan.': 'The pywin32 component needed for scanning is missing.',
 'Thu gọn bảng trang (F4)': 'Collapse page panel (F4)',
 'Thu gọn thanh công cụ (bấm đúp vào tên tab cũng được)': 'Collapse the toolbar (double-click the tab name also '
                                                          'works)',
 'Thu nhỏ': 'Zoom out',
 'Thu phóng': 'Zoom',
 'Thuộc tính máy in...': 'Printer properties...',
 'Thành phần': 'Component',
 'Thêm': 'Add',
 'Thêm bookmark': 'Add bookmark',
 'Thêm bookmark...': 'Add bookmark...',
 'Thêm chữ': 'Add text',
 'Thêm công cụ: ': 'More tools: ',
 'Thêm dòng chữ đầu trang / cuối trang. Dùng được {trang}, {tong}, {ngay}, {gio}, {ten_file}.': 'Add a header/footer '
                                                                                                'text line. You can '
                                                                                                'use {trang}, '
                                                                                                '{tong}, {ngay}, '
                                                                                                '{gio}, {ten_file}.',
 'Thêm file': 'Add file',
 'Thêm file để xem trước': 'Add a file to preview',
 'Thêm file, chọn tuỳ chọn rồi bấm Chạy': 'Add a file, choose options, then click Run',
 'Thêm ghi chú...': 'Add note...',
 'Thêm liên kết...': 'Add link...',
 'Thêm lề': 'Add margins',
 'Thêm lề trang': 'Add page margins',
 'Thêm ngày giờ và người đóng dấu:': 'Add date/time and stamped by:',
 'Thêm trang trắng': 'Add blank page',
 'Thêm vào "Open with" cho file PDF (chọn làm mặc định trong Settings > Default apps)': 'Add to "Open with" for PDF '
                                                                                        'files (set as default in '
                                                                                        'Settings > Default apps)',
 'Thêm watermark': 'Add watermark',
 'Thông tin & dữ liệu ẩn': 'Info & hidden data',
 'Thông tin & hỗ trợ': 'About & support',
 'Thông tin && hỗ trợ': 'About && support',
 'Thông tin file': 'File info',
 'Thông tin file (tác giả, phần mềm, ngày tạo...)': 'File info (author, software, creation date...)',
 'Thông tin phiên bản': 'Version info',
 'Thông tin phiên bản, liên hệ IT, giấy phép bên thứ ba': 'Version info, IT contact, third-party licenses',
 'Thông tin phiên bản, liên hệ hỗ trợ, giấy phép bên thứ ba': 'Version info, support contact, third-party licenses',
 'Thư mục': 'Folder',
 'Thư mục cài đặt': 'Install folder',
 'Thư mục cài:': 'Install folder:',
 'Thư mục không tồn tại.': 'The folder does not exist.',
 'Thư mục lưu kết quả': 'Result folder',
 'Thư mục lưu kết quả không tồn tại.': 'The result folder does not exist.',
 'Thư mục đang chọn:': 'Selected folder:',
 'Thư mục:': 'Look in:',
 "Thư viện chưa có con dấu - tạo ở tab 'Từ ảnh con dấu'.": "The library is empty - create one in the 'From stamp "
                                                           "image' tab.",
 'Thư viện con dấu': 'Stamp library',
 'Thư viện phụ thuộc': 'Dependency',
 'Thư viện đi kèm python-pptx': 'Dependency of python-pptx',
 'Thời điểm ký': 'Signing time',
 'Thứ tự': 'Order',
 'Thứ tự trang không hợp lệ.': 'Invalid page order.',
 'Thứ tự: dọc rồi sang': 'Order: down, then across',
 'Thứ tự: ngang rồi xuống': 'Order: across, then down',
 'Thử khôi phục file PDF bị hỏng, không mở được hoặc báo lỗi trong Adobe Reader.': 'Try to repair a corrupted PDF '
                                                                                   "that won't open or shows errors "
                                                                                   'in Adobe Reader.',
 'Thử lại': 'Try again',
 'Tiêu đề': 'Title',
 'Tiến': 'Forward',
 'Tiếng Anh': 'English',
 'Tiếng Việt': 'Vietnamese',
 'Tiếng Việt (đọc được cả chữ tiếng Anh xen kẽ)': 'Vietnamese (can also read mixed English text)',
 'Tiền tố': 'Prefix',
 'Token không hỗ trợ thuật toán ký SHA-256.': 'The token does not support the SHA-256 signing algorithm.',
 'Toàn bộ file': 'Entire file',
 'Toàn vẹn': 'Integrity',
 'Trang': 'Page',
 'Trang N của mẫu cho trang N': 'Template page N for page N',
 'Trang chủ': 'Home',
 'Trang cuối': 'Last page',
 'Trang cuối thành trang đầu (hay dùng cho bản scan bị ngược).': 'Last page becomes the first page (often used for '
                                                                 'reverse-scanned documents).',
 'Trang có ít hơn tỉ lệ điểm tối này được coi là trắng. Tăng nếu trang scan có nhiều chấm bẩn': 'A page with a '
                                                                                                'dark-pixel ratio '
                                                                                                'below this is '
                                                                                                'considered blank. '
                                                                                                'Increase it if '
                                                                                                'scanned pages have '
                                                                                                'a lot of '
                                                                                                'noise/specks',
 'Trang cần chuyển': 'Pages to convert',
 'Trang cần lấy': 'Pages to extract',
 'Trang cần xoá': 'Pages to delete',
 'Trang dọc': 'Portrait page',
 'Trang in': 'Pages to print',
 'Trang mẫu': 'Template page',
 'Trang mục lục': 'Table of contents page',
 'Trang ngang': 'Landscape page',
 'Trang này đang được xoay nên hình chữ ký cũng sẽ hiển thị xoay theo trang. Vẫn ký?': 'This page is rotated, so the '
                                                                                       'signature image will also '
                                                                                       'appear rotated with the '
                                                                                       'page. Sign anyway?',
 'Trang quét': 'Scanned pages',
 'Trang sau': 'Next page',
 'Trang trong file này': 'Pages in this file',
 'Trang trước': 'Previous page',
 'Trang trắng': 'Blank page',
 'Trang tuỳ chọn': 'Custom pages',
 'Trang áp dụng': 'Pages to apply',
 'Trang đã có chữ': 'Pages that already have text',
 'Trang đầu': 'First page',
 'Trang:': 'Pages:',
 'Trung bình (khuyến nghị)': 'Medium (recommended)',
 'Trung bình - ảnh 150 dpi (khuyến nghị)': 'Medium - 150 dpi images (recommended)',
 'Trên - giữa': 'Top - center',
 'Trên - phải': 'Top - right',
 'Trên - trái': 'Top - left',
 'Trình thông dịch': 'Interpreter',
 'Trích trang': 'Extract pages',
 'Trích trang...': 'Extract pages...',
 'Trích ảnh gốc': 'Extract original images',
 'Trích ảnh trong PDF': 'Extract images from PDF',
 'Trường chữ ký': 'Signature field',
 'Trắng': 'White',
 'Tuỳ chỉnh': 'Custom',
 'Tuỳ chỉnh...': 'Custom...',
 'Tuỳ chọn': 'Options',
 'Tuỳ chọn in': 'Print settings',
 'Tuỳ chọn máy in (2 mặt, khay giấy...)': 'Printer options (duplex, paper tray...)',
 'TÔ SÁNG: kéo qua đoạn chữ cần tô': 'HIGHLIGHT: drag over the text to highlight',
 'Tài liệu PDF (PDF Toolkit)': 'PDF document (PDF Toolkit)',
 'Tài liệu mới': 'New document',
 'Tác giả': 'Author',
 'Tách': 'Split',
 'Tách 1 file thành nhiều file: theo khoảng trang, mỗi N trang, từng trang, hoặc theo bookmark.': 'Split one file '
                                                                                                  'into multiple '
                                                                                                  'files: by page '
                                                                                                  'range, every N '
                                                                                                  'pages, per page, '
                                                                                                  'or by bookmark.',
 'Tách PDF': 'Split PDF',
 'Tách file...': 'Split file...',
 'Tách nền:': 'Background removal:',
 'Tách trang': 'Split pages',
 'Tách trang chọn': 'Split selected pages',
 'Tên': 'Name',
 'Tên bookmark:': 'Bookmark name:',
 'Tên con dấu:': 'Stamp name:',
 'Tên file kết quả': 'Result file name',
 'Tên file:': 'File name:',
 'Tên ô (dùng khi xuất / nhập dữ liệu form):': 'Field name (used when exporting / importing form data):',
 'Tên ô:': 'Field name:',
 'Tìm & đánh dấu xoá': 'Find & mark for redaction',
 'Tìm && thay': 'Find && Replace',
 'Tìm && thay thế': 'Find && replace',
 'Tìm && đánh dấu': 'Find && mark',
 'Tìm && đánh dấu...': 'Find && mark...',
 'Tìm chữ / số điện thoại / email / CCCD / số tài khoản để đánh dấu xoá': 'Find text / phone numbers / emails / ID '
                                                                          'numbers / account numbers to mark for '
                                                                          'redaction',
 'Tìm các trang giống hệt nhau (vd scan 2 lần) và chọn sẵn để xoá': 'Find identical pages (e.g. scanned twice) and '
                                                                    'select them for deletion',
 'Tìm công cụ...  (vd: gộp, nén, word, watermark, mật khẩu)': 'Search tools...  (e.g. merge, compress, word, '
                                                              'watermark, password)',
 'Tìm nhanh công cụ...  (vd: gộp, nén, word, mật khẩu)': 'Quick tool search...  (e.g. merge, compress, word, '
                                                         'password)',
 'Tìm trang trùng': 'Find duplicate pages',
 'Tìm trong thư mục này': 'Search this folder',
 'Tìm và thay thế': 'Find and replace',
 'Tìm và thay thế chữ': 'Find and replace text',
 'Tìm và xoá hẳn chữ / số nhạy cảm (tên, số điện thoại, email, CCCD, số tài khoản...) khỏi file - không khôi phục, không copy hay tìm lại được. Chỗ bị xoá tô đen (hoặc trắng).': 'Find '
                                                                                                                                                                                  'and '
                                                                                                                                                                                  'permanently '
                                                                                                                                                                                  'remove '
                                                                                                                                                                                  'sensitive '
                                                                                                                                                                                  'text '
                                                                                                                                                                                  '/ '
                                                                                                                                                                                  'numbers '
                                                                                                                                                                                  '(names, '
                                                                                                                                                                                  'phone '
                                                                                                                                                                                  'numbers, '
                                                                                                                                                                                  'emails, '
                                                                                                                                                                                  'ID '
                                                                                                                                                                                  'numbers, '
                                                                                                                                                                                  'account '
                                                                                                                                                                                  'numbers...) '
                                                                                                                                                                                  '- '
                                                                                                                                                                                  'they '
                                                                                                                                                                                  'cannot '
                                                                                                                                                                                  'be '
                                                                                                                                                                                  'recovered, '
                                                                                                                                                                                  'copied '
                                                                                                                                                                                  'or '
                                                                                                                                                                                  'searched. '
                                                                                                                                                                                  'Redacted '
                                                                                                                                                                                  'areas '
                                                                                                                                                                                  'are '
                                                                                                                                                                                  'filled '
                                                                                                                                                                                  'black '
                                                                                                                                                                                  '(or '
                                                                                                                                                                                  'white).',
 'Tìm:': 'Find:',
 'Tô chỗ đã xoá': 'Fill redacted areas',
 'Tô chỗ đã xoá:': 'Fill redacted areas:',
 'Tô dạ quang tự do (vùng không phải chữ)': 'Freehand highlighter (non-text areas)',
 'Tô màu nền cho trang (nằm dưới nội dung).': 'Fill the page background color (behind content).',
 'Tô nền': 'Fill background',
 'Tô nền chữ (highlight)': 'Highlight text background',
 'Tô sáng': 'Highlight',
 'Tô sáng chữ': 'Highlight text',
 'Tô sáng kèm ghi chú...': 'Highlight with note...',
 'Tô sáng vùng': 'Highlight region',
 'Tôi là người lập văn bản': "I am the document's author",
 'Tôi đồng ý với nội dung văn bản': 'I agree with the content of the document',
 'Tương phản': 'Contrast',
 'Tạo PDF từ file khác': 'Create PDF from other files',
 'Tạo booklet (in gập sách)': 'Create booklet (print and fold)',
 'Tạo bookmark theo tên từng file': "Create bookmarks from each file's name",
 'Tạo bản sao các trang đang chọn (ngay sau trang gốc)': 'Copy the selected pages (placed right after the originals)',
 'Tạo hiệu ứng bản scan': 'Make it look scanned',
 'Tạo liên kết từ địa chỉ web': 'Add links to URLs',
 'Tạo liên kết...': 'Create link...',
 'Tạo mới': 'New',
 'Tạo trang mục lục': 'Create table of contents page',
 'Tạo tài liệu PDF trống (1 trang A4) trong tab mới': 'Create a blank PDF (1 A4 page) in a new tab',
 'Tạo ô form': 'Create form field',
 'Tải lại': 'Refresh',
 'Tất cả (*.*)': 'All (*.*)',
 'Tất cả các trang': 'All pages',
 'Tất cả công cụ': 'All tools',
 'Tất cả file (*)': 'All files (*)',
 'Tất cả định dạng hỗ trợ': 'All supported formats',
 'TẨY: bấm vào ghi chú / nét vẽ / hình cần xoá': 'ERASER: click the comment / drawing / shape to delete',
 'Tẩy': 'Eraser',
 'Tắt (mặc định): dùng hộp thoại của PDF Toolkit - không bị treo khi OneDrive / Explorer chậm': 'Off (default): use '
                                                                                                "PDF Toolkit's own "
                                                                                                'dialog - does not '
                                                                                                'freeze when '
                                                                                                'OneDrive / Explorer '
                                                                                                'is slow',
 'Tệp': 'File',
 'Tỉ lệ:': 'Scale:',
 'Tối': 'Dark',
 'Tối ưu': 'Optimize',
 'Từ': 'From',
 'Từ khoá': 'Keywords',
 'Từ trái sang phải': 'Left to right',
 'Từ trên xuống dưới': 'Top to bottom',
 'Từ ảnh con dấu (tách nền)': 'From stamp image (remove background)',
 'Tự nhận diện chỗ cần điền (dòng gạch dưới / chấm chấm, đường kẻ trống, ô vuông, khung trống) và tạo ô form để người nhận gõ trực tiếp vào PDF.': 'Detect '
                                                                                                                                                   'places '
                                                                                                                                                   'to '
                                                                                                                                                   'fill '
                                                                                                                                                   'in '
                                                                                                                                                   '(underscored '
                                                                                                                                                   '/ '
                                                                                                                                                   'dotted '
                                                                                                                                                   'lines, '
                                                                                                                                                   'empty '
                                                                                                                                                   'rules, '
                                                                                                                                                   'squares, '
                                                                                                                                                   'empty '
                                                                                                                                                   'boxes) '
                                                                                                                                                   'and '
                                                                                                                                                   'create '
                                                                                                                                                   'form '
                                                                                                                                                   'fields '
                                                                                                                                                   'so '
                                                                                                                                                   'recipients '
                                                                                                                                                   'can '
                                                                                                                                                   'type '
                                                                                                                                                   'directly '
                                                                                                                                                   'into '
                                                                                                                                                   'the '
                                                                                                                                                   'PDF.',
 'Tự nhận diện chỗ trống cần điền (gạch dưới, chấm chấm, đường kẻ, ô vuông, khung trống) và tạo ô form trên mọi trang': 'Detect '
                                                                                                                        'blanks '
                                                                                                                        'to '
                                                                                                                        'fill '
                                                                                                                        'in '
                                                                                                                        '(underscores, '
                                                                                                                        'dotted '
                                                                                                                        'lines, '
                                                                                                                        'rules, '
                                                                                                                        'squares, '
                                                                                                                        'empty '
                                                                                                                        'boxes) '
                                                                                                                        'and '
                                                                                                                        'create '
                                                                                                                        'form '
                                                                                                                        'fields '
                                                                                                                        'on '
                                                                                                                        'every '
                                                                                                                        'page',
 'Tự phát hiện và xoay thẳng các trang scan bị nghiêng (tới ±5°).': 'Automatically detect and straighten skewed '
                                                                    'scanned pages (up to ±5°).',
 'Tự phát hiện và xoá trang trắng (kể cả trang scan gần như trắng).': 'Automatically detect and remove blank pages '
                                                                      '(including scanned pages that are nearly '
                                                                      'blank).',
 'Tự tạo ô form': 'Auto-create form fields',
 'Tự tạo ô form từ chỗ trống': 'Auto-create form fields from blanks',
 'Tự xoay trang cho vừa giấy': 'Auto-rotate pages to fit the paper',
 'Tự xoay trang scan bị ngược / nằm ngang': 'Auto-rotate scanned pages that are upside down / sideways',
 'Tự động - giữ màu mực con dấu (bỏ chữ đen, chữ ký, nền giấy)': 'Automatic - keep the stamp ink color (drop black '
                                                                 'text, signatures, paper)',
 'Tự động sát nội dung': 'Auto-crop to content',
 'Tự động theo ảnh': 'Automatic, based on image',
 'Viền (0 = không viền):': 'Border (0 = none):',
 'Vàng': 'Yellow',
 'Vàng nhạt': 'Light yellow',
 'Vùng chọn quá nhỏ để cắt trang.': 'The selected region is too small to crop the page.',
 'Vùng chữ ký quá nhỏ - hãy kéo một khung rộng hơn (khoảng 6 x 2 cm).': 'Signature area is too small - drag a larger '
                                                                        'box (about 6 x 2 cm).',
 "Vùng này không có chữ (có thể là ảnh scan) - dùng 'Nhận dạng chữ (OCR) trong vùng'.": 'This region has no text (it '
                                                                                        'may be a scanned image) - '
                                                                                        "use 'Recognize text (OCR) "
                                                                                        "in region'.",
 'Văn bản': 'Text',
 'Văn bản / Markdown / HTML → PDF': 'Text / Markdown / HTML → PDF',
 'Văn bản / Markdown → PDF': 'Text / Markdown → PDF',
 "Văn bản tiếng Việt nên chọn 'Tiếng Việt': trộn thêm mô hình tiếng Anh dễ làm mất dấu": 'For Vietnamese text, '
                                                                                         "choose 'Vietnamese': "
                                                                                         'adding the English model '
                                                                                         'can easily drop diacritics',
 'Vật đã ở đúng vị trí này': 'The object is already in this position',
 'Vẽ': 'Draw',
 'Về trang chủ (tất cả công cụ)': 'Back to Home (all tools)',
 'Vị trí': 'Position',
 'Vừa khung': 'Fit to frame',
 'Vừa khổ giấy': 'Fit to paper',
 'Windows 11: bấm "Show more options" hoặc Shift + chuột phải để thấy menu': 'Windows 11: click "Show more options" '
                                                                             'or Shift + right-click to see the menu',
 'Windows trên máy này không có dịch vụ quét ảnh (WIA).': 'Windows on this PC has no image scanning service (WIA).',
 'Xanh da trời': 'Sky blue',
 'Xanh dương': 'Blue',
 'Xanh lá': 'Green',
 'Xanh mực': 'Ink blue',
 'Xanh nhạt': 'Light blue',
 'Xem': 'View',
 'Xem / sửa / thêm bookmark. Mỗi dòng: cấp | tiêu đề | trang.': 'View / edit / add bookmarks. Each line: level | '
                                                                'title | page.',
 'Xem / sửa ghi chú của khối': "View / edit the object's note",
 'Xem / sửa thông tin file': 'View / edit file information',
 'Xem PDF': 'View PDF',
 'Xem các lớp (layer) trong PDF bản vẽ / thiết kế và đặt lớp nào hiện, lớp nào ẩn khi mở file.': 'View the layers of '
                                                                                                 'drawing / design '
                                                                                                 'PDFs and choose '
                                                                                                 'which are shown or '
                                                                                                 'hidden when the '
                                                                                                 'file opens.',
 'Xem danh sách file nhúng trong PDF, xoá file không cần hoặc sửa mô tả.': 'List files embedded in the PDF, remove '
                                                                           'unneeded ones or edit their '
                                                                           'descriptions.',
 'Xem dạng lưới': 'Grid view',
 'Xem dạng lưới để kéo thả sắp xếp nhiều trang': 'Grid view to drag and drop to arrange pages',
 'Xem kích thước trang': 'Page dimensions',
 'Xem trước': 'Preview',
 'Xem trước 2 trang đầu ở khung bên phải (Kết quả / File gốc). Chữ nhận dạng sai sửa trực tiếp trong trình sửa (chế độ Sửa chữ). Khoảng 3-8 giây / trang.': 'Preview '
                                                                                                                                                            'of '
                                                                                                                                                            'the '
                                                                                                                                                            'first '
                                                                                                                                                            '2 '
                                                                                                                                                            'pages '
                                                                                                                                                            'on '
                                                                                                                                                            'the '
                                                                                                                                                            'right '
                                                                                                                                                            '(Result '
                                                                                                                                                            '/ '
                                                                                                                                                            'Original). '
                                                                                                                                                            'Fix '
                                                                                                                                                            'misrecognized '
                                                                                                                                                            'text '
                                                                                                                                                            'directly '
                                                                                                                                                            'in '
                                                                                                                                                            'the '
                                                                                                                                                            'editor '
                                                                                                                                                            '(Edit '
                                                                                                                                                            'text '
                                                                                                                                                            'mode). '
                                                                                                                                                            'About '
                                                                                                                                                            '3-8 '
                                                                                                                                                            'seconds '
                                                                                                                                                            '/ '
                                                                                                                                                            'page.',
 'Xem trước 2 trang đầu ở khung bên phải (Kết quả / File gốc). Chữ nhận dạng sai sửa trực tiếp trong trình sửa (chế độ Sửa file). Khoảng 3-8 giây / trang.': 'Preview '
                                                                                                                                                             'of '
                                                                                                                                                             'the '
                                                                                                                                                             'first '
                                                                                                                                                             '2 '
                                                                                                                                                             'pages '
                                                                                                                                                             'on '
                                                                                                                                                             'the '
                                                                                                                                                             'right '
                                                                                                                                                             '(Result '
                                                                                                                                                             '/ '
                                                                                                                                                             'Original). '
                                                                                                                                                             'Fix '
                                                                                                                                                             'misrecognized '
                                                                                                                                                             'text '
                                                                                                                                                             'directly '
                                                                                                                                                             'in '
                                                                                                                                                             'the '
                                                                                                                                                             'editor '
                                                                                                                                                             '(Edit '
                                                                                                                                                             'file '
                                                                                                                                                             'mode). '
                                                                                                                                                             'About '
                                                                                                                                                             '3-8 '
                                                                                                                                                             'seconds '
                                                                                                                                                             '/ '
                                                                                                                                                             'page.',
 'Xem trước:': 'Preview:',
 'Xem tất cả ghi chú trong file, tới trang, xoá, xuất bản tổng hợp': 'See all comments in the file, go to page, '
                                                                     'delete, export a summary',
 'Xem và sửa Tiêu đề, Tác giả, Chủ đề, Từ khoá... (nạp sẵn từ file đầu tiên).': 'View and edit Title, Author, '
                                                                                'Subject, Keywords... (preloaded '
                                                                                'from the first file).',
 'Xem, sửa chữ kiểu Word, chèn ảnh / liên kết / ghi chú / hình vẽ': 'View, edit text like Word, insert images / '
                                                                    'links / notes / drawings',
 'Xem, sửa, ký số, OCR file PDF': 'View, edit, sign and OCR PDF files',
 'Xoay & xoá': 'Rotate & delete',
 'Xoay góc tuỳ ý': 'Rotate by any angle',
 'Xoay nội dung trang theo góc bất kỳ (vd 1.5°, 15°, 45°), khổ trang giữ nguyên, chữ vẫn chọn được.': 'Rotate page '
                                                                                                      'content by '
                                                                                                      'any angle '
                                                                                                      '(e.g. 1.5°, '
                                                                                                      '15°, 45°); '
                                                                                                      'page size is '
                                                                                                      'kept and text '
                                                                                                      'stays '
                                                                                                      'selectable.',
 'Xoay phải': 'Rotate right',
 'Xoay trang': 'Rotate pages',
 'Xoay trái': 'Rotate left',
 'Xoay tất cả hoặc một số trang 90° / 180° / 270°.': 'Rotate all or some pages by 90° / 180° / 270°.',
 'Xong': 'Done',
 'Xoá': 'Clear',
 'Xoá con dấu này khỏi thư viện?': 'Remove this stamp from the library?',
 'Xoá các ghi chú, tô sáng, hình vẽ, bình luận khỏi file.': 'Remove notes, highlights, drawings, and comments from '
                                                            'the file.',
 'Xoá các trang chỉ định khỏi file.': 'Delete the specified pages from the file.',
 'Xoá các trang đang chọn': 'Delete selected pages',
 'Xoá cả dữ liệu ẩn (tác giả, metadata, chữ ẩn, file đính kèm, JavaScript)': 'Also remove hidden data (author, '
                                                                             'metadata, hidden text, attachments, '
                                                                             'JavaScript)',
 'Xoá dòng chữ này': 'Delete this text line',
 'Xoá dòng này': 'Delete this line',
 'Xoá dữ liệu ẩn': 'Remove hidden data',
 'Xoá ghi chú': 'Delete note',
 'Xoá ghi chú / tô sáng': 'Delete notes / highlights',
 'Xoá hẳn mọi thứ trong các vùng đã đánh dấu - không khôi phục được sau khi lưu': 'Remove everything inside the '
                                                                                  'marked areas - cannot be '
                                                                                  'recovered after saving',
 'Xoá hết': 'Delete all',
 'Xoá hết nhãn trang (về 1, 2, 3...)': 'Remove all page labels (back to 1, 2, 3...)',
 'Xoá khỏi thư viện': 'Remove from library',
 'Xoá khối chữ này': 'Delete this text block',
 'Xoá khối này': 'Delete this object',
 'Xoá liên kết': 'Delete link',
 'Xoá mọi liên kết tới trang web / email (giữ liên kết trong file)': 'Remove every link to websites / email (keeps '
                                                                     'links within the file)',
 'Xoá mọi liên kết tới trang web / email trong file?': 'Remove every link to websites / email in the file?',
 'Xoá mục đã chọn': 'Delete selected',
 'Xoá nền trắng của ảnh': 'Remove the white background',
 'Xoá thông tin': 'Redact',
 'Xoá thông tin nhiều file': 'Redact several files',
 'Xoá thông tin tác giả, JavaScript, file nhúng, chữ ẩn, liên kết... trước khi gửi ra ngoài.': 'Remove author info, '
                                                                                               'JavaScript, embedded '
                                                                                               'files, hidden text, '
                                                                                               'links... before '
                                                                                               'sending out.',
 'Xoá trang': 'Delete page',
 'Xoá trang trắng': 'Delete blank pages',
 'Xoá tất cả file đính kèm': 'Remove all attachments',
 'Xoá tất cả liên kết web': 'Remove all web links',
 'Xoá vĩnh viễn': 'Redact permanently',
 'Xoá vĩnh viễn thông tin (redact)': 'Redact sensitive information',
 'Xoá ảnh / chữ ký / con dấu này': 'Delete this image / signature / stamp',
 'Xuất PDF và mở bằng PDF Toolkit (ký số, mật khẩu, nén...)': 'Export to PDF and open in PDF Toolkit (sign, '
                                                              'password, compress...)',
 'Xuất bản tổng hợp...': 'Export summary...',
 'Xuất danh sách ghi chú': 'Export comment list',
 'Xuất file': 'Export file',
 'Xuất file đính kèm': 'Export attachments',
 'Xuất file đính kèm ra thư mục': 'Export attachments to a folder',
 'Xuất ra': 'Export',
 'Xuất thêm file .txt': 'Also export a .txt file',
 'Xuất thêm file CSV': 'Also export a CSV file',
 'Xuất tài liệu ra PDF rồi ký số / đặt mật khẩu / nén / watermark / gộp bằng PDF Toolkit': 'Export documents to PDF, '
                                                                                           'then sign / '
                                                                                           'password-protect / '
                                                                                           'compress / watermark / '
                                                                                           'merge with PDF Toolkit',
 'Xuất từng trang thành ảnh JPG, PNG, WEBP, TIFF, BMP.': 'Export each page as a JPG, PNG, WEBP, TIFF, or BMP image.',
 'Xuất từng trang thành ảnh vector SVG (phóng to không vỡ).': 'Export each page as an SVG vector image (scales '
                                                              'without quality loss).',
 'Xuống': 'Down',
 'Xác nhận': 'Confirm',
 'Xám': 'Gray',
 'Xám nhạt': 'Light gray',
 'Xóa': 'Delete',
 'Xử lý bản scan': 'Scan cleanup',
 'Xử lý ảnh': 'Image processing',
 'Yêu cầu hỗ trợ': 'Support request',
 '[HẾT HẠN] ': '[EXPIRED] ',
 'chưa nhập khoảng trang': 'page range not entered',
 'chưa nhập quy tắc nhãn trang': 'no page label rules entered',
 'chưa nhập trang cần xoá': 'pages to delete not entered',
 'chẵn': 'even',
 'chọn thư mục...': 'select folder...',
 'cuối': 'last',
 'có': 'yes',
 'file có mật khẩu': 'password-protected file',
 "file có mật khẩu - dùng công cụ 'Mở khoá PDF' trước (hoặc nhập mật khẩu)": 'password-protected file - use the '
                                                                             "'Unlock PDF' tool first (or enter the "
                                                                             'password)',
 'file có mật khẩu - mở khoá trước khi gắn dấu thời gian': 'file is password-protected - unlock it before '
                                                           'timestamping',
 'file không có bookmark': 'file has no bookmarks',
 "file không có bookmark - dùng công cụ 'Sửa bookmark' để thêm trước": "file has no bookmarks - use the 'Edit "
                                                                       "bookmarks' tool to add some first",
 'file không có lớp nội dung (layer)': 'file has no content layers',
 'khi chương trình được phân phối, mã nguồn PDF Toolkit được cung cấp theo cùng giấy phép; liên hệ': 'when the '
                                                                                                     'program is '
                                                                                                     'distributed, '
                                                                                                     'the PDF '
                                                                                                     'Toolkit source '
                                                                                                     'code is made '
                                                                                                     'available '
                                                                                                     'under the same '
                                                                                                     'license; '
                                                                                                     'contact',
 'không có kết quả PDF để xem trước': 'no PDF result to preview',
 'không phải file Office': 'not an Office file',
 'không thể xoá tất cả các trang': 'cannot delete all pages',
 'không tìm thấy bảng': 'no table found',
 'lấy khoá': 'retrieve key',
 'lẻ': 'odd',
 'một hoặc nhiều': 'one or more',
 'phần mềm tự do theo GNU AGPL v3, không có bảo hành. Mã nguồn đầy đủ đi kèm chương trình (nút Lưu mã nguồn); mã nguồn các thư viện tải tại link chính thức trong danh sách dưới đây, hoặc liên hệ': 'free '
                                                                                                                                                                                                     'software '
                                                                                                                                                                                                     'under '
                                                                                                                                                                                                     'the '
                                                                                                                                                                                                     'GNU '
                                                                                                                                                                                                     'AGPL '
                                                                                                                                                                                                     'v3, '
                                                                                                                                                                                                     'with '
                                                                                                                                                                                                     'no '
                                                                                                                                                                                                     'warranty. '
                                                                                                                                                                                                     'The '
                                                                                                                                                                                                     'complete '
                                                                                                                                                                                                     'source '
                                                                                                                                                                                                     'code '
                                                                                                                                                                                                     'is '
                                                                                                                                                                                                     'bundled '
                                                                                                                                                                                                     'with '
                                                                                                                                                                                                     'the '
                                                                                                                                                                                                     'program '
                                                                                                                                                                                                     '(Save '
                                                                                                                                                                                                     'source '
                                                                                                                                                                                                     'button); '
                                                                                                                                                                                                     'library '
                                                                                                                                                                                                     'source '
                                                                                                                                                                                                     'code '
                                                                                                                                                                                                     'is '
                                                                                                                                                                                                     'available '
                                                                                                                                                                                                     'from '
                                                                                                                                                                                                     'the '
                                                                                                                                                                                                     'official '
                                                                                                                                                                                                     'links '
                                                                                                                                                                                                     'in '
                                                                                                                                                                                                     'the '
                                                                                                                                                                                                     'list '
                                                                                                                                                                                                     'below, '
                                                                                                                                                                                                     'or '
                                                                                                                                                                                                     'contact',
 'phần mềm tự do theo GNU AGPL v3, không có bảo hành. Mã nguồn đầy đủ đi kèm chương trình (nút Lưu mã nguồn); trong 3 năm có thể yêu cầu bản sao mã nguồn, kể cả của các thư viện, qua': 'free '
                                                                                                                                                                                         'software '
                                                                                                                                                                                         'under '
                                                                                                                                                                                         'the '
                                                                                                                                                                                         'GNU '
                                                                                                                                                                                         'AGPL '
                                                                                                                                                                                         'v3, '
                                                                                                                                                                                         'with '
                                                                                                                                                                                         'no '
                                                                                                                                                                                         'warranty. '
                                                                                                                                                                                         'The '
                                                                                                                                                                                         'complete '
                                                                                                                                                                                         'source '
                                                                                                                                                                                         'code '
                                                                                                                                                                                         'is '
                                                                                                                                                                                         'bundled '
                                                                                                                                                                                         'with '
                                                                                                                                                                                         'the '
                                                                                                                                                                                         'program '
                                                                                                                                                                                         '(Save '
                                                                                                                                                                                         'source '
                                                                                                                                                                                         'code '
                                                                                                                                                                                         'button); '
                                                                                                                                                                                         'for '
                                                                                                                                                                                         '3 '
                                                                                                                                                                                         'years '
                                                                                                                                                                                         'you '
                                                                                                                                                                                         'may '
                                                                                                                                                                                         'also '
                                                                                                                                                                                         'request '
                                                                                                                                                                                         'a '
                                                                                                                                                                                         'copy '
                                                                                                                                                                                         'of '
                                                                                                                                                                                         'the '
                                                                                                                                                                                         'source '
                                                                                                                                                                                         'code, '
                                                                                                                                                                                         'including '
                                                                                                                                                                                         'the '
                                                                                                                                                                                         'libraries, '
                                                                                                                                                                                         'from',
 'sai mật khẩu': 'wrong password',
 'thiếu dữ liệu OCR (thư mục tessdata)': 'missing OCR data (tessdata folder)',
 'thư viện Qt nằm riêng trong thư mục _internal của bản cài và có thể thay bằng phiên bản tương thích.': 'the Qt '
                                                                                                         'libraries '
                                                                                                         'are kept '
                                                                                                         'separately '
                                                                                                         'in the '
                                                                                                         '_internal '
                                                                                                         'folder of '
                                                                                                         'the '
                                                                                                         'installed '
                                                                                                         'copy and '
                                                                                                         'can be '
                                                                                                         'replaced '
                                                                                                         'with a '
                                                                                                         'compatible '
                                                                                                         'version.',
 'tím': 'purple',
 'tất cả các trang đều trắng': 'all pages are blank',
 'vd 1024 KB = 1 MB, 200 KB cho cổng dịch vụ công': 'e.g. 1024 KB = 1 MB, 200 KB for public service portals',
 'vd 2 = bỏ qua trang bìa': 'e.g. 2 = skip the cover page',
 'vd:\nNguyễn Văn A\nCông ty ABC\n0123456789': 'e.g.:\nJohn Smith\nABC Company\n0123456789',
 'vd: 1-3, 4-10, 11-  (mỗi khoảng = 1 file)': 'e.g.: 1-3, 4-10, 11-  (each range = 1 file)',
 'vd: 1-3, 5, 8-': 'e.g. 1-3, 5, 8-',
 'vd: Hà Nội': 'e.g.: Hanoi',
 'vd: Xem chi tiết tại đây': 'e.g. See details here',
 'vd: https://example.com  hoặc  ten@congty.com': 'e.g.: https://example.com  or  name@company.com',
 'vd: {ngay}, tên người duyệt... để trống = không có': "e.g.: {ngay}, approver's name... leave blank = none",
 'xanh': 'blue',
 'xanh lá': 'green',
 '{n} = số trang, {tong} = tổng số trang': '{n} = page number, {tong} = total pages',
 '© 2026 TuanPA - phần mềm tự do theo GNU AGPL v3, không có bảo hành. Chạy offline. Có dùng các thành phần mã nguồn mở - xem Giấy phép bên thứ ba.': '© '
                                                                                                                                                     '2026 '
                                                                                                                                                     'TuanPA '
                                                                                                                                                     '- '
                                                                                                                                                     'free '
                                                                                                                                                     'software '
                                                                                                                                                     'under '
                                                                                                                                                     'the '
                                                                                                                                                     'GNU '
                                                                                                                                                     'AGPL '
                                                                                                                                                     'v3, '
                                                                                                                                                     'no '
                                                                                                                                                     'warranty. '
                                                                                                                                                     'Runs '
                                                                                                                                                     'offline. '
                                                                                                                                                     'Uses '
                                                                                                                                                     'open-source '
                                                                                                                                                     'components '
                                                                                                                                                     '- '
                                                                                                                                                     'see '
                                                                                                                                                     'Third-party '
                                                                                                                                                     'licenses.',
 'Áp dụng': 'Apply',
 'Áp dụng cùng giá trị cho mọi file trong danh sách (cùng mẫu form).': 'Apply the same value to every file in the '
                                                                       'list (same form template).',
 'Ô FORM: kéo vùng đặt ô điền': 'FORM FIELD: drag the area for the field',
 'Ô chữ 1 dòng': 'Single-line text',
 'Ô chữ nhiều dòng': 'Multi-line text',
 'Ô form': 'Form field',
 'Ô form không sao chép được - dùng Tạo ô form để thêm ô mới': 'Form fields cannot be copied - use Form field to add '
                                                               'a new one',
 'Ô đánh dấu (checkbox)': 'Checkbox',
 'ăâđêôơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹĐ': 'ăâđêôơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹĐ',
 'ĐA GIÁC: bấm từng đỉnh - bấm đúp / Enter / chuột phải để kết thúc, Backspace bỏ điểm cuối': 'POLYGON: click each '
                                                                                              'vertex - double-click '
                                                                                              '/ Enter / right-click '
                                                                                              'to finish, Backspace '
                                                                                              'removes the last '
                                                                                              'point',
 'Đa giác': 'Polygon',
 'Đa giác / đám mây': 'Polygon / cloud',
 'Đang cài đặt, vui lòng chờ (1-2 phút)...': 'Installing, please wait (1-2 minutes)...',
 'Đang dừng...': 'Stopping...',
 'Đang gửi trang tới máy in...': 'Sending pages to the printer...',
 'Đang ký... (nhập mã PIN của USB Token nếu được hỏi)': 'Signing... (enter your USB Token PIN if prompted)',
 'Đang mở thư mới trong ứng dụng email...': 'Opening a new message in the email app...',
 'Đang nhận dạng chữ trong vùng...': 'Recognizing text in the region...',
 'Đang nén...': 'Compressing...',
 'Đang tạo bản xem trước...': 'Creating preview...',
 'Đang xử lý...': 'Processing...',
 'Đang đọc thư mục...': 'Reading folder...',
 'Đen': 'Black',
 'Đen trắng': 'Black & white',
 'Đen trắng (văn bản, file nhỏ nhất)': 'Black & white (text, smallest file)',
 'Điền các ô form trong PDF (ô chữ, ô chọn, danh sách). Có thể làm phẳng sau khi điền.': 'Fill in PDF form fields '
                                                                                         '(text fields, checkboxes, '
                                                                                         'lists). Can be flattened '
                                                                                         'after filling.',
 'Điền form': 'Fill form',
 'Điền form PDF': 'Fill PDF form',
 'Điều hướng': 'Navigation',
 'ĐÁM MÂY: kéo để khoanh vùng': 'CLOUD: drag to mark an area',
 'ĐÁNH DẤU XOÁ: kéo vùng cần xoá vĩnh viễn (khung đỏ) - xem lại rồi bấm Xoá vĩnh viễn': 'MARK FOR REDACTION: drag '
                                                                                        'over areas to remove '
                                                                                        'permanently (red frame) - '
                                                                                        'review, then click Redact '
                                                                                        'permanently',
 'ĐÃ DUYỆT': 'APPROVED',
 'ĐÃ KIỂM TRA': 'CHECKED',
 'ĐÃ NHẬN': 'RECEIVED',
 'ĐÃ THANH TOÁN': 'PAID',
 'ĐÍNH KÈM FILE: bấm vào vị trí đặt biểu tượng file': 'ATTACH FILE: click where the file icon goes',
 'Đám mây': 'Cloud',
 'Đánh dấu': 'Mark',
 'Đánh dấu chữ': 'Text markup',
 'Đánh dấu sẵn': 'Checked by default',
 'Đánh dấu xoá': 'Mark for redaction',
 'Đánh dấu xoá vĩnh viễn': 'Mark for permanent redaction',
 'Đánh liên tục qua các file trong danh sách': 'Number continuously across files in the list',
 'Đánh mã hồ sơ liên tục qua nhiều file (vd HS-000001, HS-000002...).': 'Assign continuous record codes across '
                                                                        'multiple files (e.g. HS-000001, '
                                                                        'HS-000002...).',
 'Đánh số Bates (hồ sơ)': 'Bates numbering (records)',
 'Đánh số trang': 'Page numbering',
 'Đánh số trang với nhiều kiểu: 1, Trang 1, Trang 1/10, - 1 -...': 'Number pages in various styles: 1, Page 1, Page '
                                                                   '1/10, - 1 -...',
 'Đã chọn các trang trùng - bấm Delete để xoá': 'Duplicate pages selected - press Delete to remove them',
 'Đã cài add-in Office (Word, Excel, PowerPoint) - mở lại Office để thấy tab PDF Toolkit': 'Installed the Office '
                                                                                           'add-in (Word, Excel, '
                                                                                           'PowerPoint) - restart '
                                                                                           'Office to see the PDF '
                                                                                           'Toolkit tab',
 'Đã cài add-in Office (Word, Excel, PowerPoint, Outlook) - mở lại Office để thấy tab / nhóm PDF Toolkit': 'Installed '
                                                                                                           'the '
                                                                                                           'Office '
                                                                                                           'add-in '
                                                                                                           '(Word, '
                                                                                                           'Excel, '
                                                                                                           'PowerPoint, '
                                                                                                           'Outlook) '
                                                                                                           '- '
                                                                                                           'restart '
                                                                                                           'Office '
                                                                                                           'to see '
                                                                                                           'the PDF '
                                                                                                           'Toolkit '
                                                                                                           'tab / '
                                                                                                           'group',
 'Đã cài xong. Mở bản đã cài ngay bây giờ?\n(Lần sau mở PDF Toolkit từ Start Menu / Desktop)': 'Installation '
                                                                                               'complete. Open the '
                                                                                               'installed version '
                                                                                               'now?\n'
                                                                                               '(Next time, open PDF '
                                                                                               'Toolkit from the '
                                                                                               'Start Menu / '
                                                                                               'Desktop)',
 'Đã căn chỉnh  (Ctrl+Z để hoàn tác)': 'Aligned  (Ctrl+Z to undo)',
 'Đã cắt - Ctrl+V để dán': 'Cut - Ctrl+V to paste',
 'Đã di chuyển  (Ctrl+Z để hoàn tác)': 'Moved  (Ctrl+Z to undo)',
 'Đã dán  (Ctrl+Z để hoàn tác)': 'Pasted  (Ctrl+Z to undo)',
 'Đã dán chữ  (Ctrl+Z để hoàn tác)': 'Text pasted  (Ctrl+Z to undo)',
 'Đã dán ảnh  (Ctrl+Z để hoàn tác)': 'Image pasted  (Ctrl+Z to undo)',
 'Đã dừng.': 'Stopped.',
 'Đã gỡ shortcut, liên kết file PDF và add-in Office.': 'Removed the shortcuts, PDF file association and Office '
                                                        'add-in.',
 'Đã gỡ shortcut, liên kết file PDF, add-in Office và phần mở rộng Explorer.': 'Removed the shortcuts, PDF file '
                                                                               'association, Office add-in and '
                                                                               'Explorer extension.',
 'Đã huỷ nhập PIN.': 'PIN entry canceled.',
 'Đã huỷ thao tác.': 'Operation canceled.',
 'Đã lưu vào thư viện con dấu': 'Saved to the stamp library',
 'Đã lưu ảnh': 'Image saved',
 'Đã lưu ảnh vào bộ nhớ tạm': 'Image copied to the clipboard',
 'Đã sao chép': 'Copied',
 'Đã sao chép - Ctrl+V để dán (sang trang / file khác cũng được)': 'Copied - Ctrl+V to paste (also onto another page '
                                                                   '/ file)',
 'Đã sao chép chữ': 'Text copied',
 'Đã sao chép chữ vào bộ nhớ tạm': 'Text copied to the clipboard',
 'Đã sao chép đường dẫn': 'Path copied',
 'Đã thêm PDF Toolkit vào "Open with" cho file .pdf': 'Added PDF Toolkit to "Open with" for .pdf files',
 'Đã thêm bookmark': 'Bookmark added',
 'Đã thêm menu chuột phải "PDF Toolkit" cho file .pdf (Windows 11: trong "Show more options")': 'Added the '
                                                                                                'right-click "PDF '
                                                                                                'Toolkit" menu for '
                                                                                                '.pdf files (Windows '
                                                                                                '11: under "Show '
                                                                                                'more options")',
 'Đã thêm ảnh thu nhỏ và khung xem trước (Preview pane) file PDF trong Explorer': 'Added PDF thumbnails and the '
                                                                                  'Preview pane in Explorer',
 'Đã tạo shortcut': 'Created shortcut',
 'Đã tạo ô form': 'Form fields created',
 'Đã xoá  (Ctrl+Z để hoàn tác)': 'Deleted  (Ctrl+Z to undo)',
 'Đã xoá dữ liệu người dùng: file gần đây, cài đặt, thư viện con dấu, cache.': 'Removed user data: recent files, '
                                                                               'settings, stamp library, cache.',
 'Đã xoá vĩnh viễn': 'Redacted permanently',
 'Đã đóng dấu - kéo để di chuyển, kéo góc dưới phải để đổi cỡ, chuột phải để xoá / nhân bản': 'Stamped - drag to '
                                                                                              'move, drag the '
                                                                                              'bottom-right corner '
                                                                                              'to resize, '
                                                                                              'right-click to delete '
                                                                                              '/ duplicate',
 'Đè lên nội dung': 'Overlay on content',
 'Đè lên trên': 'Overlay on top',
 'Đính kèm file': 'Attach file',
 'Đính kèm file vào PDF': 'Attach file to PDF',
 'Đóng': 'Close',
 'Đóng các tab bên phải': 'Close tabs to the right',
 'Đóng các tab khác': 'Close other tabs',
 'Đóng dấu': 'Stamp',
 'Đóng dấu & Đánh số': 'Stamp & Number',
 'Đóng dấu & đánh số': 'Stamp & number',
 'Đóng dấu chữ': 'Text stamp',
 'Đóng dấu hơi nghiêng (giống dấu thật)': 'Slightly tilted (like a real stamp)',
 'Đóng dấu khung chữ như ĐÃ DUYỆT, BẢN SAO, MẬT... kèm ngày (tuỳ chọn).': 'Stamp text frames like APPROVED, COPY, '
                                                                          'CONFIDENTIAL... with an optional date.',
 'Đóng dấu nhiều trang': 'Stamp many pages',
 'Đóng file đang xem': 'Close the current file',
 'Đóng gói exe': 'Exe packaging',
 'Đóng gói thành 1 file ZIP': 'Package into a single ZIP file',
 'Đóng tab': 'Close tab',
 'Đóng tab  (Ctrl+W)': 'Close tab  (Ctrl+W)',
 'Đơn vị: ': 'Unit: ',
 'ĐƯỜNG GẤP KHÚC: bấm từng điểm - bấm đúp / Enter / chuột phải để kết thúc': 'POLYLINE: click each point - '
                                                                             'double-click / Enter / right-click to '
                                                                             'finish',
 'ĐƯỜNG THẲNG: kéo để vẽ': 'LINE: drag to draw',
 'ĐƯỜNG THẲNG: kéo để vẽ - kéo viền khung đường đã vẽ để di chuyển, ô góc để đổi cỡ': 'LINE: drag to draw - drag the '
                                                                                      'frame edge of a drawn line to '
                                                                                      'move it, the corner square to '
                                                                                      'resize',
 'Đưa file đang chọn lên': 'Move selected file up',
 'Đưa file đang chọn xuống': 'Move selected file down',
 'Đưa lên trên cùng': 'Bring to front',
 'Đưa mọi trang về cùng một khổ (vd A4), nội dung tự co giãn vừa trang.': 'Resize all pages to the same size (e.g. '
                                                                          'A4), content auto-scales to fit the page.',
 'Đưa trang đang chọn lên trước': 'Move selected page forward',
 'Đưa trang đang chọn xuống sau': 'Move selected page backward',
 'Đưa xuống dưới cùng': 'Send to back',
 'Đường / mũi tên': 'Line / arrow',
 'Đường dẫn thư mục - gõ rồi Enter': 'Folder path - type and press Enter',
 'Đường gấp khúc': 'Polyline',
 'Đường thẳng': 'Line',
 'Đảo màu': 'Invert colors',
 'Đảo màu (nền tối)': 'Invert colors (dark background)',
 'Đảo màu trang (nền đen chữ trắng) để đọc ban đêm. Trang sẽ thành ảnh.': 'Invert page colors (black background, '
                                                                          'white text) for night reading. The page '
                                                                          'will become an image.',
 'Đảo ngược thứ tự trang': 'Reverse page order',
 'Đảo thứ tự': 'Reverse order',
 'Đầu / chân trang': 'Header / Footer',
 'Đầu file': 'Start of file',
 'Đầu trang - giữa': 'Header - center',
 'Đầu trang - phải': 'Header - right',
 'Đầu trang - trái': 'Header - left',
 'Đầu trang / chân trang': 'Header / Footer',
 'Đặt 2 bản cạnh nhau, tô đỏ chữ bị xoá ở bản 1 và tô xanh chữ mới thêm ở bản 2.': 'Place 2 versions side by side, '
                                                                                   'highlighting deleted text in red '
                                                                                   'in version 1 and added text in '
                                                                                   'green in version 2.',
 'Đặt mật khẩu': 'Set password',
 'Đặt mật khẩu / phân quyền': 'Set password / permissions',
 'Đặt mật khẩu mở file (mã hoá AES-256) và giới hạn in, sao chép, chỉnh sửa.': 'Set a password to open the file '
                                                                               '(AES-256 encryption) and restrict '
                                                                               'printing, copying, and editing.',
 'Đặt nhãn số trang (i, ii, A-1...)': 'Page labels (i, ii, A-1...)',
 'Đặt số trang hiển thị trong trình đọc PDF: phần mở đầu i, ii, iii; phụ lục A-1, A-2...': 'Set the page numbers '
                                                                                           'shown in PDF readers: '
                                                                                           'front matter i, ii, iii; '
                                                                                           'appendix A-1, A-2...',
 'Đặt trang của một file PDF khác (giấy tiêu đề công ty, khung mẫu...) lên trên hoặc dưới mọi trang.': 'Place pages '
                                                                                                       'from another '
                                                                                                       'PDF (company '
                                                                                                       'letterhead, '
                                                                                                       'template '
                                                                                                       'frame...) on '
                                                                                                       'top of or '
                                                                                                       'beneath '
                                                                                                       'every page.',
 'Đặt vật đang chọn vào chính giữa trang': 'Place the selected object in the center of the page',
 'Đếm': 'Count',
 'Đến': 'To',
 'ĐỀ XUẤT CHÈN CHỮ: bấm vào chỗ cần chèn rồi nhập nội dung - dấu ^ đánh dấu vị trí, nội dung xem trong Danh sách ghi chú (Acrobat: rê chuột lên dấu ^)': 'INSERT '
                                                                                                                                                         'TEXT '
                                                                                                                                                         'SUGGESTION: '
                                                                                                                                                         'click '
                                                                                                                                                         'where '
                                                                                                                                                         'text '
                                                                                                                                                         'should '
                                                                                                                                                         'go, '
                                                                                                                                                         'then '
                                                                                                                                                         'type '
                                                                                                                                                         'it '
                                                                                                                                                         '- '
                                                                                                                                                         'a '
                                                                                                                                                         '^ '
                                                                                                                                                         'mark '
                                                                                                                                                         'shows '
                                                                                                                                                         'the '
                                                                                                                                                         'spot; '
                                                                                                                                                         'the '
                                                                                                                                                         'text '
                                                                                                                                                         'appears '
                                                                                                                                                         'in '
                                                                                                                                                         'the '
                                                                                                                                                         'Comment '
                                                                                                                                                         'list '
                                                                                                                                                         '(Acrobat: '
                                                                                                                                                         'hover '
                                                                                                                                                         'over '
                                                                                                                                                         'the '
                                                                                                                                                         '^)',
 'Đề xuất chèn chữ': 'Insert text suggestion',
 'Đề xuất thay chữ': 'Suggest replacement',
 'Đề xuất thay chữ...': 'Suggest replacement...',
 'Để trống = biến vùng đã chọn (chữ / ảnh có sẵn) thành liên kết': 'Leave blank = turn the selected area (existing '
                                                                   'text / image) into a link',
 'Để trống nếu chỉ muốn giới hạn quyền': 'Leave blank if you only want to restrict permissions',
 'Địa chỉ email': 'Email address',
 'Địa chỉ web / email': 'Website / email address',
 'Địa chỉ web / email:': 'Web / email address:',
 'Định dạng': 'Format',
 'Đọc XML': 'XML parsing',
 'Đọc cấu trúc chứng thư số': 'Certificate structure parsing',
 'Đọc lại danh sách chứng thư (sau khi cắm USB Token)': 'Reload certificate list (after plugging in the USB Token)',
 'Đọc, sửa, hiển thị PDF; OCR (Tesseract tích hợp)': 'Read, edit and render PDF; OCR (built-in Tesseract)',
 'Đỏ': 'Red',
 'Đồng bộ khổ giấy': 'Unify page size',
 'Đổi bên trang chẵn / lẻ (in 2 mặt)': 'Swap sides for even / odd pages (double-sided printing)',
 'Đổi chỗ 2 trang': 'Swap 2 pages',
 'Đổi chỗ 2 trang đang chọn (chọn 2 trang ở danh sách bên trái)': 'Swap the 2 selected pages (select 2 pages in the '
                                                                  'left list)',
 'Đổi khổ giấy': 'Resize pages',
 'Đổi màu / độ dày bút vẽ': 'Change pen color / width',
 'Đổi màu chữ': 'Change text color',
 'Đổi màu nền trang': 'Change page background color',
 'Đổi màu toàn bộ chữ trong PDF (vd chữ xanh / xám nhạt → đen đậm để in rõ). Hình ảnh giữ nguyên.': 'Change the '
                                                                                                    'color of all '
                                                                                                    'text in the PDF '
                                                                                                    '(e.g. blue / '
                                                                                                    'light grey → '
                                                                                                    'solid black for '
                                                                                                    'clear '
                                                                                                    'printing). '
                                                                                                    'Images are '
                                                                                                    'unchanged.',
 'Đổi tên ô form': 'Rename form field',
 'Đổi tên ô...': 'Rename field...',
 'Độ bão hoà màu': 'Saturation',
 'Độ dày viền:': 'Border thickness:',
 'Độ nghiêng tối đa': 'Maximum skew angle',
 'Độ nghiêng tối đa cần dò': 'Maximum skew to detect',
 'Độ nhạy (% điểm tối tối đa)': 'Sensitivity (max % dark pixels)',
 'Độ nét ảnh': 'Image resolution',
 'Độ phân giải': 'Resolution',
 'Độ phân giải nhận dạng': 'Recognition resolution',
 'Độ phân giải ảnh nền': 'Background image resolution',
 'Độ phân giải:': 'Resolution:',
 'Độ rộng ảnh (% trang)': 'Image width (% of page)',
 'Độ sáng': 'Brightness',
 'Độ đậm': 'Boldness',
 'đ': 'đ',
 'đen / xám': 'black / gray',
 'để trống = tất cả; vd: 1-3, 5, 8-  hoặc  le / chan': 'leave blank = all; e.g.: 1-3, 5, 8-  or  le / chan',
 'để trống = tự đặt tên': 'leave blank = auto-generate name',
 'để trống nếu không cần': 'leave blank if not needed',
 'đỏ': 'red',
 'Ảnh': 'Image',
 'Ảnh (*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp)': 'Images (*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff '
                                                             '*.webp)',
 'Ảnh (*.png *.jpg *.jpeg *.bmp *.webp)': 'Images (*.png *.jpg *.jpeg *.bmp *.webp)',
 'Ảnh (*.png *.jpg *.jpeg *.bmp)': 'Images (*.png *.jpg *.jpeg *.bmp)',
 'Ảnh / logo': 'Image / logo',
 'Ảnh 150 dpi, JPEG 75% — phù hợp gửi email, in ấn thông thường.': '150 dpi image, JPEG 75% — suitable for emailing '
                                                                   'and regular printing.',
 'Ảnh 96 dpi, JPEG 55% — file nhỏ nhất, ảnh có thể hơi mờ.': '96 dpi image, JPEG 55% — smallest file size, image may '
                                                             'be slightly blurry.',
 'Ảnh PNG (*.png)': 'PNG image (*.png)',
 'Ảnh PNG (*.png);;Ảnh JPEG (*.jpg)': 'PNG image (*.png);;JPEG image (*.jpg)',
 'Ảnh chữ ký tay:': 'Handwritten signature image:',
 'Ảnh này có kích thước bằng 0 - không di chuyển được.': 'This image has zero size - it cannot be moved.',
 'Ảnh này nằm trong một nhóm đối tượng - không di chuyển được.': 'This image is inside an object group - it cannot '
                                                                 'be moved.',
 'Ảnh này nằm trong nhiều nhóm đối tượng khác nhau - không di chuyển được.': 'This image belongs to several object '
                                                                             'groups - it cannot be moved.',
 'Ảnh này nằm trong nhóm đối tượng dùng nhiều lần - không di chuyển riêng được.': 'This image is inside a reused '
                                                                                  'object group - it cannot be moved '
                                                                                  'on its own.',
 'Ảnh này nằm trong nhóm đối tượng được vẽ nhiều lần trên trang - không di chuyển riêng được.': 'This image is '
                                                                                                'inside an object '
                                                                                                'group drawn several '
                                                                                                'times on the page - '
                                                                                                'it cannot be moved '
                                                                                                'on its own.',
 'Ảnh nền chỉ còn logo / con dấu / kẻ bảng - 150 dpi là đủ': 'The background only keeps logos / stamps / table lines '
                                                             '- 150 dpi is enough',
 'Ảnh thu nhỏ trang': 'Page thumbnail',
 'Ảnh → PDF': 'Image → PDF',
 'Ổ mạng': 'Network drive',
 'Ổ đĩa': 'Drive',
 'Ứng dụng': 'Application',
 'Ứng dụng tạo': 'Creator application',
 '✔ Nội dung không bị sửa': '✔ Content has not been modified',
 '✔ toàn vẹn': '✔ intact',
 '✘ Nội dung đã bị sửa sau khi ký': '✘ Content was modified after signing',
 '✘ ĐÃ BỊ SỬA SAU KHI KÝ': '✘ MODIFIED AFTER SIGNING'}

PATTERNS = {'   {0} {1} {2} x {3} mm: {4} trang ({5})': '   {0} {1} {2} x {3} mm: {4} pages ({5})',
 '  (chạy thử trên {0} trang đầu)': '  (test run on the first {0} pages)',
 '  |  Đang chọn {0} trang': '  |  {0} pages selected',
 ' (2 mặt: {0} tờ giấy / bản)': ' (2-sided: {0} sheets of paper per copy)',
 ' - {0} kết quả': ' - {0} results',
 ' - đã sửa: {0}': ' - modified: {0}',
 ' x {0} bản': ' x {0} copies',
 ' ở trang: {0}': ' on page: {0}',
 ', bỏ qua {0} trang đã có chữ': ', skipped {0} pages that already have text',
 ', giữ nguyên {0} trang đã có chữ': ', kept {0} page(s) that already had text',
 ', tự xoay {0} trang': ', auto-rotated {0} pages',
 '<tr><th>Đính kèm</th><td>{0}</td></tr>': '<tr><th>Attachment</th><td>{0}</td></tr>',
 'Bản 1 - trang {0}': 'Version 1 - page {0}',
 'Bản 2 - trang {0}': 'Version 2 - page {0}',
 'Bản đã cài nằm ở {0} - hãy gỡ từ bản đó (Settings > Apps > PDF Toolkit).': 'PDF Toolkit is installed in {0} - '
                                                                             'uninstall it from there (Settings > '
                                                                             'Apps > PDF Toolkit).',
 'Các nhóm trang giống hệt nhau:\n{0}\n\nChọn sẵn {1} trang bản sao (giữ trang đầu mỗi nhóm) để bạn xem / xoá?': 'Groups '
                                                                                                                 'of '
                                                                                                                 'identical '
                                                                                                                 'pages:\n'
                                                                                                                 '{0}\n'
                                                                                                                 '\n'
                                                                                                                 'Select '
                                                                                                                 'the '
                                                                                                                 '{1} '
                                                                                                                 'duplicate '
                                                                                                                 'page(s) '
                                                                                                                 '(keeping '
                                                                                                                 'the '
                                                                                                                 'first '
                                                                                                                 'of '
                                                                                                                 'each '
                                                                                                                 'group) '
                                                                                                                 'for '
                                                                                                                 'review '
                                                                                                                 '/ '
                                                                                                                 'deletion?',
 'Các trang đang chọn ({0})': 'Selected pages ({0})',
 'Còn sót (đang được dùng): {0}': 'Left behind (in use): {0}',
 'Còn {0} vùng đánh dấu xoá chưa áp dụng - thông tin bên dưới VẪN còn trong file.\n\nXoá vĩnh viễn các vùng này rồi lưu?': '{0} '
                                                                                                                           'redaction '
                                                                                                                           'marks '
                                                                                                                           'have '
                                                                                                                           'not '
                                                                                                                           'been '
                                                                                                                           'applied '
                                                                                                                           '- '
                                                                                                                           'the '
                                                                                                                           'information '
                                                                                                                           'underneath '
                                                                                                                           'is '
                                                                                                                           'STILL '
                                                                                                                           'in '
                                                                                                                           'the '
                                                                                                                           'file.\n'
                                                                                                                           '\n'
                                                                                                                           'Redact '
                                                                                                                           'these '
                                                                                                                           'areas '
                                                                                                                           'permanently, '
                                                                                                                           'then '
                                                                                                                           'save?',
 'Công cụ này dùng đúng {0} file.': 'This tool requires exactly {0} files.',
 'Cấp bởi: {0}  |  Số serial: {1}  |  Hiệu lực: {2} - {3}': 'Issued by: {0}  |  Serial number: {1}  |  Valid: {2} - '
                                                            '{3}',
 'Cần ít nhất {0} file.': 'At least {0} files are required.',
 'Cắt trang {0} chỉ giữ lại vùng đang chọn?\n(Ctrl+Z để hoàn tác)': 'Crop page {0} to the selected region only?\n'
                                                                    '(Ctrl+Z to undo)',
 "Dòng bookmark sai định dạng: '{0}' (đúng: cấp | tiêu đề | trang)": "Bookmark line has invalid format: '{0}' "
                                                                     '(correct: level | title | page)',
 "Dòng nhãn sai định dạng: '{0}' (đúng: từ trang | kiểu | tiền tố | số bắt đầu)": "Invalid label line: '{0}' "
                                                                                  '(expected: from page | style | '
                                                                                  'prefix | start number)',
 "ERROR không có công cụ '{0}'. Có: ": "ERROR no such tool '{0}'. Available: ",
 'File gốc {0} trang': 'Original file: {0} page(s)',
 'File nguồn chỉ có {0} trang - không đủ {1} trang để thay.': 'The source file has only {0} page(s) - not enough to '
                                                              'replace {1}.',
 'File đầu vào ({0} file: {1})': 'Input files ({0} files: {1})',
 'Font gốc trong PDF: {0}  →  dùng font Windows tương ứng. Bôi đen một phần chữ để đổi định dạng riêng phần đó.': 'Original '
                                                                                                                  'font '
                                                                                                                  'in '
                                                                                                                  'PDF: '
                                                                                                                  '{0}  '
                                                                                                                  '→  '
                                                                                                                  'using '
                                                                                                                  'the '
                                                                                                                  'corresponding '
                                                                                                                  'Windows '
                                                                                                                  'font. '
                                                                                                                  'Select '
                                                                                                                  'part '
                                                                                                                  'of '
                                                                                                                  'the '
                                                                                                                  'text '
                                                                                                                  'to '
                                                                                                                  'change '
                                                                                                                  'formatting '
                                                                                                                  'for '
                                                                                                                  'just '
                                                                                                                  'that '
                                                                                                                  'part.',
 'In tất cả trang ({0})': 'Print all pages ({0})',
 'Khung không đủ chỗ - chữ đã được thu nhỏ còn {0}%': 'Not enough space in the box - text has been shrunk to {0}%',
 'Khác nhau ở {0} trang: {1}{2}  (bị xoá {3} từ, thêm {4} từ)': 'Differences on {0} pages: {1}{2}  ({3} words '
                                                                'removed, {4} words added)',
 'Không gộp được file:\n{0}': 'Could not merge files:\n{0}',
 "Không hiểu phạm vi trang '{0}' (ví dụ đúng: 1-3, 5, 8-)": "Unrecognized page range '{0}' (correct example: 1-3, 5, "
                                                            '8-)',
 'Không kiểm tra được:\n{0}': 'Could not verify:\n{0}',
 'Không ký được:\n{0}': 'Could not sign:\n{0}',
 'Không lấy được dấu thời gian từ máy chủ TSA:\n{0}\n\nKý không có dấu thời gian?': 'Could not get a timestamp from '
                                                                                    'the TSA server:\n'
                                                                                    '{0}\n'
                                                                                    '\n'
                                                                                    'Sign without a timestamp?',
 'Không mở được file:\n{0}': 'Could not open file:\n{0}',
 'Không mở được máy scan ({0}).': 'Could not open the scanner ({0}).',
 'Không mở được ứng dụng email (mã lỗi MAPI {0}).\nMáy cần cài Outlook (hoặc ứng dụng email mặc định có hỗ trợ MAPI).\n\nMở thư mục chứa file để tự đính kèm?': 'Could '
                                                                                                                                                                'not '
                                                                                                                                                                'open '
                                                                                                                                                                'the '
                                                                                                                                                                'email '
                                                                                                                                                                'app '
                                                                                                                                                                '(MAPI '
                                                                                                                                                                'error '
                                                                                                                                                                '{0}).\n'
                                                                                                                                                                'Outlook '
                                                                                                                                                                '(or '
                                                                                                                                                                'a '
                                                                                                                                                                'default '
                                                                                                                                                                'mail '
                                                                                                                                                                'app '
                                                                                                                                                                'that '
                                                                                                                                                                'supports '
                                                                                                                                                                'MAPI) '
                                                                                                                                                                'is '
                                                                                                                                                                'required.\n'
                                                                                                                                                                '\n'
                                                                                                                                                                'Open '
                                                                                                                                                                'the '
                                                                                                                                                                'folder '
                                                                                                                                                                'containing '
                                                                                                                                                                'the '
                                                                                                                                                                'file '
                                                                                                                                                                'to '
                                                                                                                                                                'attach '
                                                                                                                                                                'it '
                                                                                                                                                                'manually?',
 'Không tạo được PDF từ ảnh quét:\n{0}': 'Could not create a PDF from the scanned images:\n{0}',
 'Không tạo được file đính kèm:\n{0}': 'Could not create the attachment:\n{0}',
 'Không xem trước được: {0}': 'Could not preview: {0}',
 'Không đọc được {0}:\n{1}': 'Could not read {0}:\n{1}',
 "Kiểu nhãn không hợp lệ: '{0}' (dùng 1, i, I, a, A hoặc -)": "Invalid label style: '{0}' (use 1, i, I, a, A or -)",
 'Kích thước hiện tại: <b>{0}</b>': 'Current size: <b>{0}</b>',
 'Kết quả {0} trang · {1} KB': 'Result {0} pages · {1} KB',
 'Kết quả {0} trang.': 'Result: {0} pages.',
 'Liên kết mở file:\n{0}\n\nMở?': 'The link opens a file:\n{0}\n\nOpen it?',
 'Lỗi hiển thị trang: {0}': 'Error displaying page: {0}',
 'Lỗi khi in:\n{0}': 'Print error:\n{0}',
 'Lỗi khi ký:\n{0}': 'Error while signing:\n{0}',
 'Lỗi khi lưu:\n{0}': 'Error while saving:\n{0}',
 'Lỗi khi nén:\n{0}': 'Error while compressing:\n{0}',
 'Lỗi khi quét: {0}': 'Scan error: {0}',
 'Lỗi ký số ({0}): 0x{1}': 'Digital signature error ({0}): 0x{1}',
 'Một số file đang được dùng (Office đang mở?) - xoá tay thư mục {0} sau.': 'Some files are in use (is Office open?) '
                                                                            '- delete the folder {0} manually later.',
 'Mở liên kết này?\n\n{0}': 'Open this link?\n\n{0}',
 'Nhập mật khẩu để mở:\n{0}': 'Enter password to open:\n{0}',
 "Phạm vi trang '{0}' không có trang nào (file có {1} trang)": "Page range '{0}' contains no pages (file has {1} "
                                                               'pages)',
 'Sẽ in {0} trang trên {1} tờ': '{0} pages will be printed on {1} sheets',
 'Sẽ in {0} trang. Trang ngang được xoay cho vừa khổ giấy.': '{0} page(s) will be printed. Landscape pages are '
                                                             'rotated to fit the paper.',
 'Thay "{0}" bằng:': 'Replace "{0}" with:',
 'Thay {0} trang đang chọn bằng các trang của file mới,\nbắt đầu từ trang số:': 'Replace the {0} selected page(s) '
                                                                                'with pages of the new file,\n'
                                                                                'starting at page:',
 'Thư mục {0} sẽ được xoá sau khi đóng chương trình.': 'The folder {0} will be deleted after the program closes.',
 'Trang hiện tại ({0})': 'Current page ({0})',
 'Trang {0}': 'Page {0}',
 'Trang {0}/{1}': 'Page {0}/{1}',
 'Trước: {0}\nSau:   {1}\nGiảm {2}%\n\nBấm Lưu / Lưu thành để ghi ra file.': 'Before: {0}\n'
                                                                             'After:   {1}\n'
                                                                             'Reduced by {2}%\n'
                                                                             '\n'
                                                                             'Click Save / Save As to write to file.',
 'Tách {0} trang ra file mới': 'Split {0} pages into a new file',
 'Tìm thấy {0} chỗ': 'Found {0} matches',
 'Tất cả ({0} trang)': 'All ({0} pages)',
 'Tất cả công cụ ({0})': 'All tools ({0})',
 'Xoá vĩnh viễn {0} vùng đã đánh dấu?\n\nChữ, ảnh và hình vẽ trong vùng bị xoá hẳn khỏi file (không copy / tìm lại được). Trước khi lưu vẫn hoàn tác được bằng Ctrl+Z.': 'Permanently '
                                                                                                                                                                         'redact '
                                                                                                                                                                         '{0} '
                                                                                                                                                                         'marked '
                                                                                                                                                                         'areas?\n'
                                                                                                                                                                         '\n'
                                                                                                                                                                         'Text, '
                                                                                                                                                                         'images '
                                                                                                                                                                         'and '
                                                                                                                                                                         'graphics '
                                                                                                                                                                         'inside '
                                                                                                                                                                         'are '
                                                                                                                                                                         'removed '
                                                                                                                                                                         'from '
                                                                                                                                                                         'the '
                                                                                                                                                                         'file '
                                                                                                                                                                         '(cannot '
                                                                                                                                                                         'be '
                                                                                                                                                                         'copied '
                                                                                                                                                                         'or '
                                                                                                                                                                         'searched). '
                                                                                                                                                                         'You '
                                                                                                                                                                         'can '
                                                                                                                                                                         'still '
                                                                                                                                                                         'undo '
                                                                                                                                                                         'with '
                                                                                                                                                                         'Ctrl+Z '
                                                                                                                                                                         'before '
                                                                                                                                                                         'saving.',
 'Xoá {0} trang: {1}{2}?': 'Delete {0} pages: {1}{2}?',
 'không lấy được dấu thời gian từ máy chủ ({0})': 'could not get a timestamp from the server ({0})',
 'máy chưa cài Microsoft {0}': 'Microsoft {0} is not installed on this computer',
 'tối thiểu {0}': 'minimum {0}',
 '{0} - trang {1}': '{0} - page {1}',
 '{0} - đã nhận dạng {1}/{2} trang': '{0} - recognized {1}/{2} pages',
 '{0} công cụ  |  Chạy offline - file không rời khỏi máy': '{0} tools  |  Runs offline - files never leave your '
                                                           'computer',
 '{0} file': '{0} file(s)',
 '{0} thư mục, {1} file': '{0} folder(s), {1} file(s)',
 '{0} trang / tờ': '{0} pages per sheet',
 '{0} tính năng': '{0} features',
 '{0} vùng - bấm Lưu để ghi ra file': '{0} areas - click Save to write the file',
 '{0} vùng đang chờ xoá - bấm Xoá vĩnh viễn để áp dụng': '{0} areas waiting - click Redact permanently to apply',
 '{0} ô chữ, {1} ô đánh dấu - Ctrl+Z để hoàn tác': '{0} text fields, {1} check boxes - Ctrl+Z to undo',
 '{0}: dựng lại {1} trang thành chữ sửa được ({2} từ)': '{0}: rebuilt {1} page(s) as editable text ({2} words)',
 '{0}: file không có mật khẩu / giới hạn': '{0}: file has no password / restrictions',
 '{0}: file đã tối ưu, không nhỏ hơn được': '{0}: file is already optimized, cannot be made smaller',
 '{0}: không có chữ - có thể là PDF scan': '{0}: no text - may be a scanned PDF',
 '{0}: không có chữ ký số': '{0}: no digital signature',
 '{0}: không có file đính kèm': '{0}: no attached files',
 '{0}: không thấy bảng - xuất toàn bộ chữ': '{0}: no table found - exporting all text',
 '{0}: không thấy chỗ trống để tạo ô form': '{0}: no blanks found to create form fields',
 '{0}: không tìm thấy chữ cần xoá (file scan cần OCR trước)': '{0}: no matching text (scanned files need OCR first)',
 '{0}: làm thẳng {1}/{2} trang': '{0}: straightened {1}/{2} pages',
 '{0}: nhận dạng {1} trang': '{0}: recognized {1} pages',
 '{0}: thử mức {1}': '{0}: trying level {1}',
 '{0}: tạo {1} ô ({2} ô chữ, {3} ô đánh dấu)': '{0}: created {1} fields ({2} text fields, {3} check boxes)',
 '{0}: xoá {1} file, còn {2} file': '{0}: removed {1} files, {2} remaining',
 '{0}: xoá {1} ghi chú': '{0}: removed {1} notes',
 '{0}: xoá {1} trang trắng': '{0}: removed {1} blank pages',
 '{0}: {1} KB → {2} KB (giảm {3}%)': '{0}: {1} KB → {2} KB (reduced by {3}%)',
 '{0}: {1} bảng': '{0}: {1} tables',
 '{0}: {1} lớp hiện, {2} lớp ẩn': '{0}: {1} layers shown, {2} hidden',
 '{0}: {1} trang': '{0}: {1} pages',
 '{0}: {1} trang → 1 trang {2} x {3} mm': '{0}: {1} pages → 1 page {2} x {3} mm',
 '{0}: {1} trang → {2} mặt in ({3} tờ)': '{0}: {1} pages → {2} printed sides ({3} sheets)',
 '{0}: {1} ảnh': '{0}: {1} images',
 '{0}: điền {1} ô': '{0}: filled {1} fields',
 '{0}: đã gắn dấu thời gian ({1})': '{0}: timestamp added ({1})',
 '{0}: đã xoá {1} chỗ': '{0}: redacted {1} matches',
 '{0}{1}  (CA: {2}, hết hạn {3})': '{0}{1}  (CA: {2}, expires {3})',
 'Ô form: {0}': 'Form field: {0}',
 'Đang chạy bản không cài (bản đã cài: {0})': 'Running a non-installed copy (installed copy: {0})',
 'Đang cài bản {0} - cài lại để cập nhật lên bản {1}.': 'Version {0} is installed - install again to update to '
                                                        'version {1}.',
 'Đi tới trang {0}': 'Go to page {0}',
 'Đã bỏ {0} vùng đánh dấu xoá': 'Removed {0} redaction marks',
 'Đã chèn {0} trang quét sau trang {1}': 'Inserted {0} scanned pages after page {1}',
 'Đã chèn {0} trang từ {1} file': 'Inserted {0} pages from {1} files',
 'Đã chép chương trình vào {0}': 'Copied the program to {0}',
 'Đã chọn {0} ký tự - Ctrl+C để sao chép, chuột phải để đánh dấu': '{0} character(s) selected - Ctrl+C to copy, '
                                                                   'right-click to mark up',
 'Đã chọn {0} từ - Ctrl+C để sao chép, chuột phải để đánh dấu': '{0} word(s) selected - Ctrl+C to copy, right-click '
                                                                'to mark up',
 'Đã chụp vùng {0}x{1} px vào clipboard - dán (Ctrl+V) vào Word, Outlook, Zalo...': 'Captured {0}x{1} px to the '
                                                                                    'clipboard - paste (Ctrl+V) into '
                                                                                    'Word, Outlook, Zalo...',
 'Đã cài vào máy: {0}': 'Installed on this PC: {0}',
 "Đã có ô form tên '{0}'.": "A form field named '{0}' already exists.",
 'Đã gộp {0} file, {1} trang.': 'Merged {0} files, {1} pages.',
 'Đã gửi {0} trang tới máy in {1}': 'Sent {0} page(s) to printer {1}',
 'Đã ký số thành công.\n\nNgười ký: {0}\nFile: {1}\n\nLưu ý: không sửa / lưu đè file này nữa, nếu không chữ ký sẽ mất hiệu lực.': 'Digitally '
                                                                                                                                  'signed '
                                                                                                                                  'successfully.\n'
                                                                                                                                  '\n'
                                                                                                                                  'Signer: '
                                                                                                                                  '{0}\n'
                                                                                                                                  'File: '
                                                                                                                                  '{1}\n'
                                                                                                                                  '\n'
                                                                                                                                  'Note: '
                                                                                                                                  'do '
                                                                                                                                  'not '
                                                                                                                                  'edit '
                                                                                                                                  '/ '
                                                                                                                                  'overwrite '
                                                                                                                                  'this '
                                                                                                                                  'file '
                                                                                                                                  'anymore, '
                                                                                                                                  'or '
                                                                                                                                  'the '
                                                                                                                                  'signature '
                                                                                                                                  'will '
                                                                                                                                  'become '
                                                                                                                                  'invalid.',
 'Đã lưu {0} ({1})': 'Saved {0} ({1})',
 'Đã lưu {0} trang vào {1}': 'Saved {0} pages to {1}',
 'Đã nhân bản {0} trang  (Ctrl+Z để hoàn tác)': 'Duplicated {0} page(s)  (Ctrl+Z to undo)',
 'Đã quét {0} trang - bấm Lưu để ghi ra file, OCR để nhận dạng chữ': 'Scanned {0} pages - click Save to write the '
                                                                     'file, OCR to recognize text',
 'Đã quét {0} trang.\n\nĐặt trang tiếp theo lên mặt kính rồi bấm Yes để quét tiếp, No để kết thúc.': 'Scanned {0} '
                                                                                                     'pages.\n'
                                                                                                     '\n'
                                                                                                     'Place the next '
                                                                                                     'page on the '
                                                                                                     'glass and '
                                                                                                     'click Yes to '
                                                                                                     'continue, or '
                                                                                                     'No to finish.',
 'Đã thay {0} chỗ.': 'Replaced {0} occurrences.',
 'Đã thay {0} trang  (Ctrl+Z để hoàn tác)': 'Replaced {0} page(s)  (Ctrl+Z to undo)',
 'Đã thêm ảnh thu nhỏ trang PDF trong Explorer (khung xem trước giữ của máy: {0})': 'Added PDF thumbnails in '
                                                                                    "Explorer (keeping this PC's "
                                                                                    'preview pane: {0})',
 'Đã tạo {0} liên kết từ địa chỉ web / email trong file.': 'Created {0} link(s) from web / email addresses in the '
                                                           'file.',
 "Đã tạo ô form '{0}'": "Form field '{0}' created",
 'Đã xoá {0}': 'Deleted {0}',
 'Đã xoá {0} liên kết web': 'Removed {0} web link(s)',
 'Đã xuất {0}': 'Exported {0}',
 'Đã áp dụng: {0}  (Ctrl+Z để hoàn tác)': 'Applied: {0}  (Ctrl+Z to undo)',
 'Đã đánh dấu {0} chỗ trên {1} trang (khung đỏ) - xem lại rồi bấm Xoá vĩnh viễn': 'Marked {0} matches on {1} pages '
                                                                                  '(red frame) - review, then click '
                                                                                  'Redact permanently',
 'đúng {0}': 'exactly {0}',
 '⚠ {0}: chỉ nén được còn {1} KB (chưa đạt {2} KB) - file nhiều chữ / vector khó nén hơn': '⚠ {0}: could only '
                                                                                           'compress to {1} KB '
                                                                                           '(target {2} KB not '
                                                                                           'reached) - files with '
                                                                                           'lots of text / vectors '
                                                                                           'are harder to compress',
 '✔ Xong - {0} kết quả': '✔ Done - {0} results'}

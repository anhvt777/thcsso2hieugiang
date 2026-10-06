# THCS Số 2 Hiếu Giang · SchoolCollect

Ứng dụng web tĩnh để theo dõi sĩ số, khoản phải thu và kết quả đối soát. Giao diện tối ưu cho máy tính kế toán và hiệu trưởng, đồng thời dùng được trên điện thoại.

## Quyền riêng tư

- Đọc file ngay trong trình duyệt; không có API gửi file lên máy chủ và không dùng analytics/CDN.
- Học sinh, khoản thu, giao dịch và lịch sử được lưu trong IndexedDB của trình duyệt trên thiết bị đang sử dụng.
- Các máy không tự đồng bộ với nhau. Có thể chuyển dữ liệu bằng tệp sao lưu `.sctbackup` được mã hóa AES-GCM bằng mật khẩu.
- GitHub Pages chỉ lưu mã nguồn và giao diện. Không commit danh sách học sinh, sao kê hoặc tệp sao lưu có dữ liệu thật.
- Có thể dùng ngoại tuyến sau lần mở web đầu tiên khi service worker đã lưu giao diện.

## Nhập dữ liệu

- Hỗ trợ `.xlsx`, `.csv`, `.tsv` và `.txt`; trong Excel nhiều trang tính, tự chọn trang có tiêu đề phù hợp và nhiều dòng dữ liệu nhất. Không tải thư viện từ CDN.
- Danh sách theo mẫu trường có hai dòng cho mỗi học sinh. Web gộp theo “Mã học sinh”, lưu riêng mã “Mã HS theo Khoản nộp”, loại khoản BHYT/BHTT và số tiền; kiểm tra trùng mã và thiếu khoản trước khi lưu.
- Báo cáo ngân hàng: nhận diện “Mã khách hàng”, “Số hóa đơn”, “Ngày giao dịch”, “Tên khách hàng”, “Số tiền” và “Trạng thái giao dịch”. Mã khách hàng đuôi YT được ghép với BHYT, đuôi TT được ghép với BHTT. Trạng thái không thành công không được tính đã thu.
- Dịch vụ khác có thể xem chi tiết theo nội dung chuyển khoản như gửi xe, nước uống. Mã học sinh không tìm thấy sẽ hiện là chưa khớp.
- Một giao dịch chỉ được tính đã thu khi mã khách hàng/mã khoản, loại khoản, trạng thái thành công và số tiền khớp chính xác với một món phải thu chưa được ghép. Giao dịch sai số tiền, sai khoản, trùng món hoặc không tìm thấy học sinh không cộng vào số đã thu.
- Các mẫu CSV/XLSX cũ với mã học sinh và cột khoản thu vẫn được hỗ trợ qua bước ghép cột.
- Tạo mã QR: lưu một lần mã BIN, số tài khoản và tên chủ tài khoản của trường; sau đó tạo ảnh QR riêng cho từng món phải thu, lọc theo lớp và tải ZIP. Nội dung thanh toán được tạo thành ảnh QR và đóng gói ngay trong trình duyệt bằng thư viện cục bộ.
- QR mặc định chỉ tạo cho món chưa có giao dịch khớp chính xác. Sau khi cập nhật báo cáo thu, danh sách QR được tính lại. Hãy quét thử bằng ứng dụng ngân hàng trước khi gửi cho phụ huynh.

## Triển khai

Mở qua máy chủ web tĩnh HTTPS, chẳng hạn GitHub Pages. Tránh mở trực tiếp bằng `file://` vì trình duyệt có thể chặn IndexedDB và service worker. Với GitHub Pages của repository, đặt `index.html` ở thư mục gốc của nhánh được chọn làm nguồn Pages.

Trước khi dùng dữ liệu thật, hãy đối chiếu tên cột và cách ghi mã học sinh trong file gốc của trường/ngân hàng; kiểm tra tổng số dòng và tổng tiền với báo cáo nguồn.

## QR Bảo hiểm · 2 phương án

- Mở tab **QR Bảo hiểm · 2 phương án**, chọn chính xác khoản BHYT và BHTT đã phân giao, năm học, hạn nộp và lớp. Không tự gán mức đóng mới.
- Kiểm tra và lưu tài khoản của trường ở mục **Tài khoản nhận tiền**, dùng chung cấu hình QR hiện có. Xem trước rồi xuất ZIP gồm PNG từng học sinh, PDF A5 từng lớp và bảng mã CSV. Tên file gồm BHYT, BHTT và lớp.
- Chưa nộp: có QR BHYT và QR BHYT + BHTT. Đã nộp một khoản: chỉ còn QR khoản còn lại. Đã hoàn thành: không có QR, có thể xuất thông báo hoàn thành bằng bộ lọc “Tất cả”. Hồ sơ thiếu hoặc trùng khoản bị loại và được liệt kê để kiểm tra.
- Đây là QR chuyển khoản đến tài khoản đã cấu hình, có mã tham chiếu `IB` + 20 ký tự hex. Web **không tự đăng ký mã khách hàng, hóa đơn thu hộ hoặc tài khoản định danh tại BIDV**. Mã được tạo ổn định theo học sinh, khoản thu, số tiền, tài khoản và năm học; không cắt ngắn mã học sinh để tạo mã.
- Để tự đối soát QR này, báo cáo ngân hàng cần có mã IB đầy đủ trong nội dung chuyển khoản, mã thanh toán hoặc tham chiếu. Báo cáo thu hộ chỉ có mã KH cũ và không có mã IB sẽ không đủ để xác định phương án mới. Cần kiểm tra luồng nhận báo cáo phù hợp trước khi phát hành thật.
- Mã phương án được lưu cùng bảng phân bổ từng khoản trong `meta.noticeBundles` và bản sao lưu mã hóa. Không xóa mã đã phát hành khi xuất lại. Sao lưu sau khi tạo thông báo, khôi phục bản sao lưu nếu đổi máy; CSV tra cứu không thay thế bản sao lưu.
- Giao dịch có mã IB chỉ khớp khi mã, khoản và số tiền khớp nguyên vẹn. Không dùng tổng tiền để suy đoán phương án; sai tiền, mã lạ, thiếu khoản hoặc mã mâu thuẫn được giữ lại để kiểm tra. Thanh toán cả hai QR: giao dịch chồng khoản bị đánh dấu trùng, không tự cộng lần hai. Kế toán cần xác minh và xử lý tiền thừa; web không tự hoàn tiền.
- QR ảnh đã gửi **không tự hết hiệu lực và không khóa được tại ngân hàng**. Cập nhật báo cáo trước khi xuất nhắc thu; thông báo yêu cầu phụ huynh chỉ trả một phương án và không trả lại khoản đã nộp.
- QR có vùng trắng bốn module, không chèn logo vào mã. Trước khi phát hành toàn trường, quét thử bằng ứng dụng ngân hàng để kiểm tra người nhận, số tiền và mã nội dung; chưa thực hiện chuyển tiền chỉ để thử giao diện.

Kiểm thử logic bằng dữ liệu giả: `node tests/insurance.test.cjs`.

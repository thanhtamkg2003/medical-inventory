# Kiểm kê thiết bị y tế bằng QR

## Chạy trên máy tính Windows

Giải nén bộ phần mềm, mở Command Prompt tại thư mục chứa `app.py`:

```bat
py -m pip install -r requirements.txt
py app.py
```

Mở `http://127.0.0.1:5056/login`. Tài khoản khởi tạo: `admin` / `admin123`.
Đăng nhập admin, vào menu tên đăng nhập → **Tài khoản** → **Khởi tạo Superadmin**.
Nhập mật khẩu admin để xác nhận, đặt tên và mật khẩu riêng cho Superadmin,
sau đó ghi lại mã khôi phục chỉ hiển thị một lần. Hãy cất mã ở nơi an toàn,
đổi mật khẩu mặc định của admin ngay.

Khi nâng cấp từ bản trước, **thay `app.py`, `requirements.txt` và tùy chọn `README.md`; giữ nguyên `medical_inventory.db`**. Bản mới tự bổ sung các bảng/cột cần thiết cho nhật ký, phiên đăng nhập và đồng bộ offline; dữ liệu danh mục, tài khoản, các đợt và kết quả quét cũ được giữ nguyên. Hãy sao lưu `medical_inventory.db` trước khi cập nhật.

## Tài khoản, phân quyền và phiên đăng nhập

Superadmin có thể thêm admin/thành viên, cấp hoặc thu hồi quyền, khóa/mở khóa,
đặt lại mật khẩu cho tài khoản khác, xem nhật ký hệ thống và **sao lưu cơ sở dữ liệu**.
Trang Tổng quan của Superadmin có thêm các số liệu quản trị và các hoạt động gần đây. Superadmin không thể
bị xóa/khóa/cấp lại mật khẩu từ tài khoản khác. Người có quyền quản lý tài khoản
chỉ quản lý thành viên; không thể cấp quyền hoặc thao tác với admin/Superadmin.

Phiên đăng nhập có hai giới hạn:

- **8 giờ không hoạt động**: phiên tự hết hạn.
- **24 giờ tối đa**: dù người dùng vẫn hoạt động liên tục, phiên cũng phải đăng nhập lại.

Một tài khoản **có thể đăng nhập trên nhiều máy cùng lúc**. Mỗi phiên được lưu riêng
với thời gian đăng nhập, hoạt động gần nhất, IP và thông tin trình duyệt/thiết bị.
Superadmin có thể yêu cầu **đăng xuất tất cả phiên** của một tài khoản.

Khi đổi mật khẩu, thay đổi quyền hoặc khóa tài khoản, các phiên cũ bị vô hiệu hóa.
Lịch sử tên tài khoản đã quét vẫn được giữ trong báo cáo.

## Nhật ký hệ thống

Trang **Nhật ký** chỉ dành cho Superadmin. Có hai nhóm chính:

- **Lịch sử đăng nhập**: đăng nhập thành công/thất bại, đăng xuất và các phiên đang hoạt động.
- **Nhật ký hoạt động**: thêm/sửa/xóa thiết bị, nhập Excel, tạo/xóa đợt kiểm kê,
  quét QR, tạo tài khoản, đổi quyền, khóa/mở khóa, đổi/đặt lại mật khẩu, xuất báo cáo,
  xuất QR, từ chối truy cập và các thao tác quản trị quan trọng khác.

Mỗi bản ghi có thể lưu thời gian, tài khoản, vai trò, hoạt động, đối tượng,
kết quả, IP, vị trí địa lý **ước tính**, ISP/tổ chức mạng, trình duyệt/hệ điều hành
và nội dung chi tiết. Mật khẩu thật không được ghi vào nhật ký.

Có thể lọc theo tài khoản, loại nhật ký, thành công/thất bại, khoảng thời gian
và từ khóa; có thể xuất toàn bộ kết quả lọc ra Excel.

### Vị trí IP

Phần mềm ghi IP mà máy chủ nhận được. Khi IP là địa chỉ Internet công cộng,
phần mềm có thể tra cứu thành phố/khu vực/quốc gia, ISP và tọa độ gần đúng qua
dịch vụ IP geolocation. Đây là **vị trí ước tính**, không phải vị trí chính xác của
máy hay số nhà.

Trong mạng LAN, IP thường là dạng `192.168.x.x`, `10.x.x.x` hoặc tương tự; các IP nội bộ
không được gán vị trí Internet và sẽ hiển thị là **Không xác định (IP nội bộ)**.

Mặc định bản này dùng endpoint HTTPS của `ipapi.co`. Nếu dịch vụ không khả dụng,
đăng nhập vẫn tiếp tục bình thường và vị trí sẽ để trạng thái chưa xác định.
Có thể thay endpoint bằng biến môi trường `IP_GEO_API`.

Nếu ứng dụng đặt sau reverse proxy, chỉ bật `TRUST_PROXY=1` khi reverse proxy của bạn
được tin cậy để chuyển tiếp `X-Forwarded-For` đúng cách.

## Dữ liệu và nhập Excel

Trong **Thiết bị → Thêm từ danh sách Excel**, nhập sheet `TỔNG TÀI SẢN`:
cột A–K lần lượt là STT, MÃ TÀI SẢN, TÊN VT - TB, SỐ SERI, MODEL, HÃNG SX,
NƯỚC SX, NĂM SX, NĂM SD, NĂM BC TĂNG, KHOA SD. Tiêu đề ở dòng 7,
dòng 8 là chú thích, dữ liệu từ dòng 9. Có thể tải file mẫu trong trang Thiết bị.
Giao diện danh mục giữ các cột chính; báo cáo Excel giữ đủ A–K.

- Khi danh mục trống: nhập nguyên danh sách, kể cả những dòng thiếu hoặc trùng
  seri để dễ rà soát. Seri thiếu/trùng không tạo QR và không được điểm danh.
- **Cập nhật danh mục**: chỉ thêm seri chưa có. Dòng trùng seri hiện hành hoặc
  trùng với một dòng đã nhận trong file tải lên sẽ bị bỏ qua. Dòng thiếu seri
  được thêm nếu mã tài sản chưa có dòng thiếu seri, đồng thời được báo để sửa.
- **Thay danh mục**: sau xác nhận, danh mục hiện hành chỉ còn các dòng của file
  mới. Đợt đang làm được đóng; tạo đợt mới trước khi quét tiếp. Danh mục và
  kết quả từng đợt cũ vẫn lưu để xuất báo cáo lịch sử.

Trang kết quả nhập có số lượng và danh sách chi tiết cần rà soát; tải Excel để
xem tất cả hoặc xuất riêng các dòng thiếu seri, trùng seri của lần nhập đó.
Trang Thiết bị cũng có hai nút xuất danh sách thiếu seri và trùng seri của
danh mục hiện hành, giữ các cột A–K và thông tin đối chiếu.

Trong tất cả các bảng dữ liệu, hàng tiêu đề được **căn giữa ngang và dọc** để
giao diện đồng nhất; nội dung dữ liệu vẫn giữ cách căn phù hợp với loại thông tin.

## QR, kiểm kê và báo cáo

Mỗi tem QR chỉ mang nội dung seri và in ba dòng `Seri: ...`, `MTS: ...`, `Khoa: ...`
phía dưới. Trang Thiết bị có hai nút xuất Word: mã đã chọn trên trang hoặc
toàn bộ danh mục. Khổ A4 dọc, tối đa 20 tem/trang (4 cột × 5 hàng). Các tem
thiếu/trùng seri bị bỏ qua.

Tạo **Đợt kiểm kê** rồi chọn **Quét QR**. Mỗi seri chỉ ghi một lần trong một
đợt. Khi camera hoặc mã QR không đọc được, mở mục “Không quét được mã QR?” để
nhập seri thủ công; hệ thống báo lỗi nếu seri không có trong danh sách.

### Quét khi mất kết nối

Bản nâng cấp có **hàng đợi offline trên trình duyệt bằng IndexedDB**.
Nếu QR đã được camera đọc nhưng máy không liên lạc được với Flask server:

1. Hệ thống **không báo đã kiểm kê thành công**.
2. Mã được lưu tạm trên thiết bị và hiển thị trạng thái đang chờ đồng bộ.
3. Người dùng có thể tiếp tục quét các mã khác.
4. Khi kết nối trở lại, hệ thống tự động đồng bộ; có thêm nút **Đồng bộ ngay**.
5. Mỗi lần quét có một mã sự kiện riêng để tránh ghi trùng nếu yêu cầu đồng bộ bị gửi lại.
6. Nếu máy chủ từ chối một mã do seri không còn hợp lệ hoặc đã được ghi nhận ở nơi khác,
   mã đó được giữ trong hàng đợi ở trạng thái lỗi để người dùng rà soát, thay vì âm thầm mất dữ liệu.

Hàng đợi được lưu bền vững trên trình duyệt, nhưng nếu người dùng xóa dữ liệu
trình duyệt/IndexedDB thủ công thì các mã chưa đồng bộ có thể mất. Vì vậy nên
đồng bộ trước khi đổi máy hoặc xóa dữ liệu trình duyệt.

Báo cáo Excel gồm A–K, kết quả, thời gian và tài khoản quét theo từng đợt.
Trang Kết quả lọc được Tất cả / Có / Chưa điểm danh, kết hợp tìm kiếm và phân trang.

## Bảo mật và sao lưu

Bản nâng cấp dùng CSRF token cho các thao tác thay đổi dữ liệu; mật khẩu chỉ được
lưu dưới dạng hash. Phiên đăng nhập dùng cookie HttpOnly/SameSite và có giới hạn
8 giờ không hoạt động, 24 giờ tối đa.

Tài khoản bị vô hiệu hóa không bị xóa vật lý khỏi cơ sở dữ liệu để lịch sử hoạt động
và nhật ký vẫn giữ được danh tính tài khoản. Những thao tác xóa thiết bị/danh mục
vẫn dùng trạng thái lịch sử như các bản trước.

Superadmin có thể vào **menu tài khoản → Sao lưu dữ liệu** để tạo và tải một bản sao
SQLite nhất quán của toàn bộ `medical_inventory.db`. Thao tác sao lưu cũng được ghi vào nhật ký.
Ngoài ra, vẫn nên tự lưu thêm một bản sao bên ngoài máy trước các thao tác lớn như
**Thay danh mục**, **Xóa toàn bộ danh mục**, xóa đợt hoặc nâng cấp phần mềm.

Trước khi đưa hệ thống lên Internet:

- đổi mật khẩu khởi tạo;
- tạo và đặt `APP_SECRET` cố định, đủ dài;
- dùng máy chủ WSGI phù hợp thay cho `app.run()` thử nghiệm;
- triển khai HTTPS, đặc biệt khi quét QR bằng điện thoại;
- chỉ bật `TRUST_PROXY=1` khi reverse proxy được cấu hình tin cậy;
- cân nhắc dùng dịch vụ IP geolocation có kế hoạch sử dụng phù hợp thay cho endpoint thử nghiệm mặc định.

Bản chạy thử tại `127.0.0.1` chủ yếu dành cho máy tính cục bộ.

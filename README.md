# XLCellHighlight
 Tô màu nổi bật ô đang chọn với Formatting Condtitions sử dụng Hàm UDF VBA


Hàm tạo định dạng có điều kiện để tô màu cột và dòng làm nổi bật để đối chiếu ô đang chọn, giúp chúng ta có cái nhìn trực quan trong bảng tính Excel.

Mặc dù có thể tạo định dạng có điều kiện bằng Trình tạo định dạng có điều kiện có sẵn của Excel. Tuy nhiên để thuận tiện, làm cho việc tạo dễ dàng và nhanh chóng hơn, chỉ cần tận dụng VBA để tự động hóa chúng. Các hàm dưới đây chỉ là bổ trợ tạo Formatting conditions, chứ không phải hàm chức năng.

#### Ưu điểm của cách tô màu nổi bật với định dạng có điều kiện:
- Đối chiếu dòng cột dễ dàng.
- Định dạng không làm ảnh hướng đến chế độ Undo và Redo của Excel.
- Dễ dàng tạo định dạng điều kiện với việc gõ hàm.
#### Nhược điểm:
- Cách tô màu với định dạng có điều kiện gây tốn kém tài nguyên, khi thực hiện tính toán lại, nếu vùng ô quá lớn.

***Nếu dự án của bạn đã có nhiều công thức tính toán thì không nên sử dụng cách này để tô màu.

Hình ảnh làm nổi bật ô với định dạng có điều kiện:

![Highlight cell activated](https://github.com/user-attachments/assets/bcc19366-3063-4c9c-b81c-c4a077355585)

Hướng dẫn:
Hàm:
```=CellHighlight(RangeEvent,[Đối_số_cài_đặt])```

Gõ 1 hàm duy nhất cho 1 vùng ô cần tô màu, vào một ô bất kỳ không sử dụng đến.

Các hàm đối số cài đặt

​
| Hàm đối số cài đặt |	Diễn giải
| ------------------ |	---------------
| HL_Column(color=0) |		Nhập hàm thì tô cột với màu chỉ định hoặc màu mặc định
| HL_Row(color=0) |		Nhập hàm thì tô dòng với màu chỉ định hoặc màu mặc định
| HL_ActiveCell(color=0, borderColor=0)	 |	Nhập hàm thì tô ô chọn với màu chỉ định hoặc màu mặc định
| CellHighlight_ShowWindow() |		Nhập hàm này ô bất kì để mở Cửa sổ Formatting Conditions
| CellHighlight_Delete() |		Xóa điều kiện đã đặt, hàm có sẵn trong ô, thêm _Delete và nhấn Enter. Hoặc nếu nhập hàm này trên vùng ô đang có định dạng.
| CellHighlight_DeleteAll() |		Xóa tất cả điều kiện tô màu
| CellHighlight_CopyCode() |		Chép mã vào bộ nhớ tạm để dán vào mã ThisWorkbook
| CellHighlight_HuongDan() |		Tự động tạo trang tính hướng dẫn sử dụng


Ví dụ: Tô màu vùng ô A1:Z1000 được đặt tên là Table1, với tô cột, dòng và ô chọn\

```=CellHighlight(Table1,HL_Row(),HL_Column(),HL_ActiveCell())```\

Ví dụ: Tô màu vùng ô A1:Z1000 được đặt tên là Table1, tô dòng màu #5CC8B3 và ô chọn màu mặc định\

```=CellHighlight(Table1,HL_Row("#5CC8B3"),HL_ActiveCell())```\

Ví dụ CellHighlight_Delete:\
Nếu ô chứa hàm ```=CellHighlight(Table1,HL_ActiveCell())``` để xóa chỉ cần thêm _Delete và nhấn Enter\
Nếu nhập =CellHighlight_Delete() trên vùng ô có định dạng thì thực hiện xóa.

#### Chọn màu sắc cho định dạng màu nền:

Để đặt màu sắc có thể chọn màu trong bảng chọn màu, hoặc một số đại diện màu sắc hoặc tên màu tiếng Anh.​
​
Dưới bảng chọn màu hãy chọn màu Hệ thập lục phân, ví dụ chọn và nhập HL_Row("#5CC8B3")​

![color Picker](https://github.com/user-attachments/assets/8a92de01-f01e-4227-9114-45be8ba2e67f)

Nếu bạn đã tạo xong các định dạng có điều kiện cho các vùng ô, bạn có thể xóa tất cả mã trong module đi, để lại các dòng mã mà mã trong ThisWorkbook gọi.

Để sử dụng được Hàm trong dự án mới, hãy sao chép module modXLCellHighlight và mã trong ThisWorkbook.
***Lưu ý: khi sử dụng mã VBA thì dự án của bạn cần lưu ở các định dạng xlsm, xlsb hoặc dạng add-in xla, xlam.

Các bạn có thể tải Add-in, cài đặt để sử dụng lại cho các dự án khác


Hàm hỗ trợ thêm trong việc tô màu nền trực quan với Add-in XLCellHighlight xlam:

Ví dụ bạn muốn tô màu từ ô F4 cho đến F4:J25, hãy nhập màu nền cho F4 trắng và J4 cam, thường thì là màu sáng và nhập công thức sau:​
```=FillColor(F4:J25)​```
​
Nếu bạn muốn chọn mô hình màu, thì gồm các hàm dưới đây, với 2 tham số, vị trí màu, và khoảng cách không gian màu:​
​
```Fill_HUE(Starting, Fractor)​```\
```Fill_Natural(Starting, Fractor)​```\
```Fill_Lightness(Starting, Fractor)​```\
​
Công thức nhập như sau: ```=FillColor(F4:J25, Fill_Natural(0, 20))​```\
​
Starting luôn bắt đầu từ 0, nếu bạn muốn lệch bao nhiêu thì thêm vào. Khoảng cách không gian màu tùy màu mà đối số có thể âm.​
Các hàm này phải được nhập trong hàm FillColor.​
​
Sau khi tô xong bạn có thể xóa hàm đã gõ đi.​
​
Hình ảnh kết quả​

![Model color](https://github.com/user-attachments/assets/a0b1f5be-9f3c-4ec4-8c8b-00c085445128)

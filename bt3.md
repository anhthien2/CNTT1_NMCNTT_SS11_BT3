# Phần 1 - Phân tích 4 tham số VLOOKUP trong bài Don_Hang:

**1. Lookup_value:** Là mã linh kiện cần tìm, chính là ô A2 trong Sheet Don_Hang. Nên bọc thêm TRIM(A2) để cắt khoảng trắng thừa.

**2. Table_array:** Là toàn bộ bảng tra cứu nằm ở Sheet Danh_Muc, vùng A2:C101. Phải cố định tuyệt đối thành $A$2:$C$101 để khi kéo công thức xuống không bị lệch.

**3. Col_index_num:** Là số thứ tự cột muốn lấy. Lấy Tên linh kiện thì điền 2, lấy Đơn giá thì điền 3.

**4. Range_lookup:** Điền FALSE để yêu cầu tìm kiếm chính xác tuyệt đối mã linh kiện.


# PHẦN 2 - TRIỂN KHAI CÔNG THỨC HOÀN CHỈNH


**1. Để lấy Tên Linh Kiện (Cột B - Don_Hang):**  
Đặt tại ô B2 của Sheet Don_Hang:

excel

```
=(VLOOKUP(TRIM(A2), Danh_Muc!$A$2:$C$101, 2, FALSE), "Mã không tồn tại")
```

**2. Để lấy Đơn Giá (Cột C - Don_Hang):**  
Đặt tại ô C2:

excel

```
=(VLOOKUP(TRIM(A2), Danh_Muc!$A$2:$C$101, 3, FALSE), 0)
```



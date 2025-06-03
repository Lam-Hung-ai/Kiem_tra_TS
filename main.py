import pypandoc
import os

def convert_markdown_to_word(markdown_file_path, output_word_path):
    """
    Chuyển đổi tệp Markdown sang tệp Word (.docx) sử dụng pandoc,
    với sự hỗ trợ cho công thức toán học.

    Args:
        markdown_file_path (str): Đường dẫn đến tệp Markdown đầu vào.
        output_word_path (str): Đường dẫn để lưu tệp Word đầu ra.

    Returns:
        bool: True nếu chuyển đổi thành công, False nếu có lỗi.
    """
    try:
        # Kiểm tra xem tệp Markdown có tồn tại không
        if not os.path.exists(markdown_file_path):
            print(f"Lỗi: Không tìm thấy tệp Markdown tại '{markdown_file_path}'")
            return False

        # Sử dụng pypandoc để gọi pandoc
        # Các tùy chọn 'extra_args' có thể được sử dụng để tinh chỉnh quá trình chuyển đổi
        # Ví dụ, '--mathml' hoặc '--mathjax' có thể được xem xét tùy thuộc vào cách bạn muốn
        # công thức được render trong Word. Mặc định pandoc thường làm tốt việc này.
        # Đối với công thức toán học, pandoc thường tự động xử lý chúng dưới dạng đối tượng OMML (Office Math Markup Language) trong Word.
        pypandoc.convert_file(
            markdown_file_path,
            'docx',
            outputfile=output_word_path,
            extra_args=['--standalone'] # Đảm bảo tài liệu là một tệp hoàn chỉnh
        )
        print(f"Chuyển đổi thành công! Tệp Word đã được lưu tại: '{output_word_path}'")
        return True
    except Exception as e:
        print(f"Đã xảy ra lỗi trong quá trình chuyển đổi: {e}")
        print("Hãy đảm bảo rằng bạn đã cài đặt pandoc và thêm nó vào PATH hệ thống.")
        print("Bạn có thể tải pandoc từ: https://pandoc.org/installing.html")
        return False

# --- Ví dụ sử dụng ---
if __name__ == "__main__":
    # Thay đổi các đường dẫn này cho phù hợp với tệp của bạn
    input_md_file = "/home/lamhung/code/markdown2docx/Kiem_tra_TS/chat-export.md"  # Đường dẫn đến tệp Markdown của bạn
    output_docx_file = "tmp.docx" # Tên tệp Word bạn muốn xuất ra

    # Thực hiện chuyển đổi
    convert_markdown_to_word(input_md_file, output_docx_file)

    # (Tùy chọn) Xóa tệp Markdown mẫu sau khi chạy
    # if os.path.exists(input_md_file) and input_md_file == "thuc_nghiem_HMM.md":
    #     os.remove(input_md_file)
    #     print(f"Đã xóa tệp Markdown mẫu: '{input_md_file}'")
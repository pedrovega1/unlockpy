import itertools
import win32com.client as win32

def try_password(doc_path, password):
    word = win32.Dispatch("Word.Application")
    try:
        doc = word.Documents.Open(doc_path, PasswordDocument=password)
        doc.Close(SaveChanges=False)
        return True
    except Exception as e:
        print(f"Ошибка при попытке с паролем {password}: {e}")
        return False
    finally:
        word.Quit()

def brute_force(doc_path):
    for password_tuple in itertools.product('0123456789', repeat=6):
        password = ''.join(password_tuple)
        if try_password(doc_path, password):
            print(f"Пароль найден: {password}")
            return
    print("Пароль не найден.")

if __name__ == "__main__":
    doc_path = r"" #путь к файлу
    brute_force(doc_path)

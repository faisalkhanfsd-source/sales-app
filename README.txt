چلانے کا طریقہ (اپنے کمپیوٹر پر):
  pip install -r requirements.txt
  set APP_USER=آپ_کا_نام  &  set APP_PASSWORD=آپ_کا_پاسورڈ   (Linux/Mac: export)
  python app.py      ->  http://localhost:5000

آن لائن ہوسٹ کرنے کے لیے (Render / Railway):
  فولڈر GitHub پر ڈالیں، Build: pip install -r requirements.txt
  Start: gunicorn app:app --timeout 300 --workers 1
  Environment میں APP_USER, APP_PASSWORD, SECRET_KEY سیٹ کریں۔

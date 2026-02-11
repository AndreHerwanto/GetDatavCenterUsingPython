# 1. Install prerequisites
sudo apt update
sudo apt install python3-venv python3-full -y

# 2. Buat virtual environment
cd /home/linux/getdata
python3 -m venv venv

# 3. Aktifkan
source venv/bin/activate

# 4. Install dependencies
pip install pyvmomi urllib3

# 5. Aktifkan dulu venv-nya
source venv/bin/activate

# 6. Jalankan script
python vcenter_export_fixed.py

#7. Akan Muncul Hasil Get Data nya Format CSV
<img width="1087" height="592" alt="GetDataSample" src="https://github.com/user-attachments/assets/412c4e84-acb4-44ff-b876-3eba3bec88d9" />

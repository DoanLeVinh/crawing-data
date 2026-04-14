# Deploy Len Render

## Cach 1: Deploy bang render.yaml (khuyen nghi)

1. Day code len GitHub.
2. Vao Render -> New + -> Blueprint.
3. Chon repository chua project nay.
4. Render se tu doc file `render.yaml` va tao web service.
5. Bam `Apply` de deploy.

## Cau hinh da co san
- Build command: `sh -c 'if [ -f requirements.txt ]; then pip install -r requirements.txt; elif [ -f HoaDon/requirements.txt ]; then pip install -r HoaDon/requirements.txt; else echo "requirements.txt not found"; exit 1; fi'`
- Start command: `sh -c 'if [ -f app.py ]; then gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120; elif [ -f HoaDon/app.py ]; then gunicorn --chdir HoaDon app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120; else echo "app.py not found"; exit 1; fi'`
- Health check: `/`
- Python runtime: `python-3.11.9`

## Cach 2: Tao service thu cong
Neu khong dung Blueprint, tao Web Service voi:
- Runtime: Python
- Build command: `sh -c 'if [ -f requirements.txt ]; then pip install -r requirements.txt; elif [ -f HoaDon/requirements.txt ]; then pip install -r HoaDon/requirements.txt; else echo "requirements.txt not found"; exit 1; fi'`
- Start command: `sh -c 'if [ -f app.py ]; then gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120; elif [ -f HoaDon/app.py ]; then gunicorn --chdir HoaDon app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120; else echo "app.py not found"; exit 1; fi'`

## Fix loi "Could not open requirements file"
Neu log bao khong tim thay `requirements.txt`, nghia la Render dang build o sai folder.

Ban co 2 cach:
1. Trong Render service settings, dat **Root Directory** = `HoaDon`.
2. Hoac giu root mac dinh, dung build/start command resilient nhu o tren (tu dong thu ca `./` va `./HoaDon`).

## Luu y
- Du lieu file CSV hien dang nam trong repo, service se doc truc tiep khi khoi dong.
- Render free plan co the sleep khi khong co traffic.

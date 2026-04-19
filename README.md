# CSA ICoFR Dashboard — PT PNM

Dashboard monitoring hasil Control Self-Assessment (CSA) Internal Control over Financial Reporting (ICoFR) PT Permodalan Nasional Madani.

## Cara Deploy ke Vercel

### Opsi A: Deploy via GitHub (Recommended)

1. **Buat repo GitHub baru**
   - Buka https://github.com/new
   - Buat repository baru (bisa private)

2. **Push project ini ke GitHub**
   ```bash
   cd csa-dashboard
   git init
   git add .
   git commit -m "Initial commit - CSA Dashboard"
   git branch -M main
   git remote add origin https://github.com/USERNAME/REPO-NAME.git
   git push -u origin main
   ```

3. **Connect ke Vercel**
   - Buka https://vercel.com dan login/signup
   - Klik "Add New Project"
   - Import repository GitHub yang barusan dibuat
   - Framework Preset: pilih **Vite**
   - Klik **Deploy**
   - Tunggu ~1-2 menit, selesai!

### Opsi B: Deploy via Vercel CLI

1. **Install Vercel CLI**
   ```bash
   npm install -g vercel
   ```

2. **Login ke Vercel**
   ```bash
   vercel login
   ```

3. **Deploy**
   ```bash
   cd csa-dashboard
   npm install
   vercel
   ```
   - Ikuti prompt-nya, pilih defaults semua
   - Untuk production deploy: `vercel --prod`

## Cara Run Lokal

```bash
cd csa-dashboard
npm install
npm run dev
```

Buka http://localhost:5173

## Stack

- React 18
- Vite 5
- Recharts (charts)
- Deploy: Vercel

# 🚀 Pasos para publicar Centro Finanzas en GitHub Pages

## 1. Crea el repositorio en GitHub
- Ve a https://github.com/new
- Nombre: `centro-finanzas`
- Visibilidad: Public (necesario para GitHub Pages gratis)
- NO inicialices con README (ya lo tienes)
- Clic en "Create repository"

## 2. En tu PC, abre CMD en la carpeta del proyecto y ejecuta:

```bash
git init
git add .
git commit -m "🚀 Initial commit - Centro Conciencia Financiera"
git branch -M main
git remote add origin https://github.com/JohnBe99/centro-finanzas.git
git push -u origin main
```

## 3. Activa GitHub Pages + GitHub Actions
- Ve a tu repo → Settings → Pages
- Source: selecciona "GitHub Actions"
- Guarda

## 4. El primer deploy tarda ~2 minutos
- Ve a la pestaña "Actions" en tu repo para ver el progreso
- Cuando termine, tu app estará en:
  👉 https://JohnBe99.github.io/centro-finanzas

## 5. Para actualizaciones futuras (cada vez que edites el código):
```bash
git add .
git commit -m "descripción del cambio"
git push
```
Y GitHub Actions despliega automáticamente en ~1 minuto.

## Instalación local (para probar antes de publicar):
```bash
npm install
npm start
```
Abre http://localhost:3000

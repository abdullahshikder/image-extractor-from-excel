# Publishing to GitHub

Your repository has been initialized and committed locally. To make it public on GitHub, follow these steps:

## Step 1: Create a GitHub Repository

1. Go to [GitHub](https://github.com) and sign in
2. Click the "+" icon in the top right corner
3. Select "New repository"
4. Repository name: `excel-image-extractor`
5. Description: "Extract images from Excel files and rename them with product names"
6. Choose **Public** visibility
7. **DO NOT** initialize with README, .gitignore, or license (we already have these)
8. Click "Create repository"

## Step 2: Push to GitHub

After creating the repository, GitHub will show you commands. Run these in your terminal:

```bash
cd "C:\Users\Pathao Ltd.PATHAO-LTD\Documents\image extractor"
git remote add origin https://github.com/YOUR_USERNAME/excel-image-extractor.git
git branch -M main
git push -u origin main
```

Replace `YOUR_USERNAME` with your actual GitHub username.

## Alternative: Using SSH

If you prefer SSH:

```bash
git remote add origin git@github.com:YOUR_USERNAME/excel-image-extractor.git
git branch -M main
git push -u origin main
```

## Verify

After pushing, visit your repository URL:
`https://github.com/YOUR_USERNAME/excel-image-extractor`

Your repository should now be public and accessible to everyone!


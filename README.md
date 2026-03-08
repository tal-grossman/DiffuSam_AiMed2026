# DiffuSam_AiMed2026

AIMed 2026 abstract and submission materials for **DiffuSAM: Diffusion-Based Prompt-Free Medical Image Segmentation via SAM2 Memory Embeddings**.

This repo is intended to be used as a **submodule** of the main [sam2_mem_diffusion](https://github.com/...) project (or your fork).

## Contents

- `aimed2026_abstract.tex` — LaTeX abstract (compile with `pdflatex aimed2026_abstract.tex`)
- `AIMed2026-submission-template.docx` — Conference submission template
- Pipeline, t-SNE, and inference-step figures (PNG)

## Using a remote (GitHub/GitLab)

1. Create an empty repo named `DiffuSam_AiMed2026` on your host.
2. In this repo:
   ```bash
   git remote add origin https://github.com/YOUR_USER/DiffuSam_AiMed2026.git
   git push -u origin master
   ```
3. In the **parent** repo (`sam2_mem_diffusion`), point the submodule at the remote:
   ```bash
   git config submodule.AiMed2026.url https://github.com/YOUR_USER/DiffuSam_AiMed2026.git
   git submodule sync
   ```

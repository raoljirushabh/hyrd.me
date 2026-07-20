# Hyrd.me

Tailor your resume, cover letter & application — naturally, strategically, undetectably.

**Hyrd** is an AI-powered tool that helps job seekers optimize their application materials for maximum ATS (Applicant Tracking System) compatibility while maintaining authenticity and human-like writing quality.

## Features

- **Resume Tailoring**: Automatically refine your resume to match job descriptions with strategic keyword placement
- **Cover Letter Generation**: Create personalized cover letters that stand out
- **Application Q&A**: Generate thoughtful answers to common application questions
- **ATS Scoring**: Get detailed ATS compatibility scores and breakdowns
- **File Support**: Upload PDF or DOCX resume files for parsing
- **Export Options**: Download results as .txt or .docx files
- **Dark Mode UI**: Professional, modern dark theme interface

## Tech Stack

- **React 18** - UI framework
- **TypeScript** - Type safety
- **Vite** - Fast build tool
- **Mammoth** - DOCX file parsing
- **PDF.js** - PDF file parsing
- **Claude AI** - LLM for content generation

## Getting Started

### Prerequisites

- Node.js 16+ 
- npm or yarn

### Installation

```bash
npm install
```

### Development

```bash
npm run dev
```

The app will open at `http://localhost:3000`

### Build

```bash
npm run build
```

## Usage

1. **Paste Job Description**: Copy and paste the full job posting
2. **Upload Your Resume**: Either paste your resume text or upload a PDF/DOCX file
3. **Add Application Questions** (Optional): Include any specific questions from the job application
4. **Generate**: Click to tailor your materials
5. **Download**: Export your results as text or Word documents

## API Requirements

This app uses the Anthropic Claude API. Set up your API key in the application to use the AI features.

## License

MIT

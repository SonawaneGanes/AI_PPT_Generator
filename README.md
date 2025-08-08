# AI_PPT_Generator

The AI-Powered Presentation Generator is a Flask-based web application that leverages the OpenRouter GPT-3.5 API to create professional, fully-themed PowerPoint presentations instantly. By entering a topic, users can generate a .pptx file complete with custom slide layouts, color themes, typography, and relevant images — ready for download and use in meetings, lectures, or projects.

This project is designed to save time and effort in presentation creation by automating the slide-building process. It integrates AI text generation with a dynamic design system, ensuring that both the content and the visual appeal of the slides meet professional standards.

✨ Features
AI-Generated Content – Automatically creates slide titles and bullet points from any user-provided topic.

Dynamic Theme Selection – Choose from multiple pre-defined color and font themes to match the style of your presentation.

Automated Image Fetching – Retrieves high-quality, royalty-free images from Unsplash based on the slide topic.

Custom Slide Layouts – Designed with clear typography, sidebar accents, and modern color palettes for professional appeal.

Responsive Web Interface – Built with HTML, CSS, and JavaScript for a smooth and intuitive user experience.

Secure API Management – Uses .env file for storing API keys securely.

Downloadable PPTX – Generates and sends the final .pptx file directly to the user’s device.

🛠 Tech Stack
Backend: Python, Flask

Frontend: HTML5, CSS3, JavaScript

AI Integration: OpenRouter GPT-3.5 API

Presentation Engine: python-pptx

Image Source: Unsplash API

🚀 How It Works
User enters a topic into the web interface.

Flask backend sends the topic to OpenRouter GPT-3.5 for content generation.

python-pptx formats the generated text into slides, applying the selected theme.

Relevant images are fetched from Unsplash and inserted into the slides.

The final presentation is packaged as a .pptx file and made available for download.

📌 Use Cases
Business presentations

Academic lectures

Startup pitch decks

Event overviews

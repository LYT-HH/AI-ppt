# AI-ppt
📋 Project Overview
AI PPT Generator is an intelligent presentation generation tool powered by artificial intelligence technology that can instantly transform user ideas into professional-grade slides. This project combines DeepSeek large language models, python-pptx library, and Streamlit framework to achieve full-process automation from topic input to PPT file generation.
✨ Core Features
1. Intelligent Content Generation

AI-Driven Outline Generation: Simply input a topic, and AI automatically generates a complete PPT outline including cover page, table of contents structure, core content for each section, and summary page.
Structured Content Output: Based on DeepSeek API, generates professional content suitable for various scenarios including business reports, academic presentations, product introductions, and more.
Multi-Style Adaptation: Supports multiple styles including business reporting, academic speeches, product introductions, project proposals, and educational training.

2. Custom Template Support

Template Upload Function: Users can upload their own PPT template files (.pptx format), and the system will generate PPTs based on user templates.
Smart Placeholder Filling: Automatically identifies title and content areas in templates, filling AI-generated content into corresponding placeholders.
Default Template System: Built-in various professional templates to meet different scenario requirements, ensuring high-quality PPT generation even without custom templates.

3. Data Visualization Integration

Automatic Chart Generation: Automatically adds data visualization elements such as bar charts, line charts, and pie charts based on content.
Intelligent Layout Optimization: Automatically adjusts layout and design style to ensure visual dynamic balance.
Real-time Preview Editing: Supports online custom editing, allowing modification of text content, styles, images, and other elements.

🛠️ Technical Architecture
Backend Technology Stack

AI Engine: DeepSeek API (compatible with OpenAI format), used for content generation and structured processing
PPT Processing: python-pptx library, supporting PPT file creation, editing, and template application
Web Framework: Streamlit, providing interactive web interface and file upload functionality
Asynchronous Processing: Supports multi-round dialogue state management, implementing rollback-capable workflows

Frontend Features

Responsive Interface: Modern web interface based on Streamlit
Real-time Progress Display: Status feedback and progress bars during generation process
File Management: Supports template upload, generation result download, and local storage

🚀 Quick Start
Environment Requirements

Python 3.8+
DeepSeek API Key

Installation Steps
bash# Clone the project
git clone https://github.com/yourusername/ai-ppt-generator.git
cd ai-ppt-generator

# Install dependencies
pip install -r requirements.txt

# Configure API key
# Set your DeepSeek API key in the code

Run the Application
bashstreamlit run ai_ppt_generator.py

Visit http://localhost:8501 to start using.
📖 Usage Guide
1. Basic Usage Process

Input Topic: Enter PPT topic in the web interface, such as "Future Development of Artificial Intelligence"
Select Settings: Set parameters like number of slides, style preferences, etc.
Upload Template (Optional): Upload custom PPT template file
Generate PPT: Click the generate button and wait for AI processing to complete
Download Results: Download the generated .pptx file for further editing

2. Template Usage Tips

Template Preparation: It is recommended to design template files with clear placeholders in PowerPoint
Compatibility: Supports all standard .pptx format templates, ensuring templates contain title and content placeholders
Brand Consistency: Maintain corporate brand style consistency through custom templates

3. Content Optimization Suggestions

Clear Topic: Provide specific, clear topic descriptions for more accurate AI-generated content
Style Matching: Choose appropriate PPT style based on usage scenario
Post-Editing: Generated PPTs can be further beautified and adjusted in PowerPoint or WPS

🔧 Configuration Instructions
API Configuration
python# Configure DeepSeek API in the code
DEEPSEEK_API_KEY = "your_api_key_here"
DEEPSEEK_BASE_URL = "https://api.deepseek.com"

Template Configuration

Default template path: default_template.pptx
User template save directory: user_templates/
Supported template format: .pptx (standard PowerPoint format)

📊 Performance Characteristics

Generation Speed: Typically 10-30 seconds to complete generation of an 8-page PPT
Content Quality: Based on DeepSeek large models, content is professional and accurate
Template Compatibility: Supports custom templates, maintaining brand consistency
Output Format: Standard .pptx format, compatible with Microsoft PowerPoint and WPS

🎯 Application Scenarios
This project is suitable for various office and educational scenarios:

Business Reporting: Enterprise quarterly reports, project proposals, market analysis
Academic Presentations: Thesis defense, academic exchanges, research achievement displays
Education and Training: Teaching courseware, training materials, knowledge sharing
Product Introduction: Product launches, brand displays, marketing materials

🔄 Development Roadmap
Implemented Features

 Basic AI content generation
 Custom template support
 Data visualization charts
 Web interaction interface

Planned Features

 Multi-model support (integrating more AI models)
 Batch generation functionality
 Team collaboration support
 Mobile adaptation

🤝 Contribution Guidelines
Welcome to submit Issues and Pull Requests to help improve the project. Before submitting code, please ensure:

Code complies with PEP8 specifications
Add appropriate test cases
Update relevant documentation

📄 License
This project uses the MIT License. See LICENSE file for details.
📞 Support and Feedback
For questions or suggestions, please contact through:

GitHub Issues: Submit problem reports
Email: Project maintainer email
Documentation: View project Wiki for more information

🌟 Acknowledgments
Thanks to the following open-source projects and services for support:

DeepSeek AI for content generation capabilities
python-pptx library for PPT processing functionality
Streamlit framework for web interface support
All contributors and users for feedback and support

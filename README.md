# AI-Powered PPT Generator

This project is a web application that generates PowerPoint presentations based on user-provided topics and the number of slides. It uses the OpenAI API to generate slide content and the `python-pptx` library to create the PowerPoint files.

## Project Structure
- `app.py`: The main Flask application file.
- `requirements.txt`: Lists the dependencies required for the project.
- `static/`: Contains static files such as CSS and generated presentations.
- `templates/`: Contains HTML templates for the web pages.
- `uploads/`: Contains uploaded and generated PowerPoint files.

## Requirements

- Python 3.x
- Flask
- Requests
- python-pptx
- OpenAI
- python-dotenv

## Installation

1. Clone the repository:
    ```sh
    git clone <repository-url>
    cd <repository-directory>
    ```

2. Create a virtual environment and activate it:
    ```sh
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3. Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

4. Set up your OpenAI API key:
    - Create a `.env` file in the root directory of the project.
    - Add your OpenAI API key to the `.env` file:
      ```
      OPENAI_API_KEY=your_openai_api_key
      ```

## Usage

1. Run the Flask application:
    ```sh
    python app.py
    ```

2. Open your web browser and go to `http://127.0.0.1:5000`.

3. Enter the topic and the number of slides you want in the form and click "Generate & Download".

4. The generated PowerPoint presentation will be downloaded automatically.
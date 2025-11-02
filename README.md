# Random PowerPoint Generator

A vibecoded Python script that generates random educational PowerPoint presentations on the fly. Every line of code and this description written by AI. Perfect for practicing your presentation skills by forcing you to present on topics you've never researched before - true zero-shot presentation practice! 🎯

## Features

- 🎲 Random topic generation within a specified category
- 📊 Auto-generated slide content with detailed explanations
- 🖼️ AI-generated images (DALL-E 3) for each content slide
- 📐 Professional 16:9 aspect ratio presentations
- 🎨 Properly formatted with titles, bullets, and hierarchical content
- ⚡ Quick setup and execution

## Prerequisites

- Python 3.x
- OpenAI API key ([get one here](https://platform.openai.com/api-keys))
- Virtual environment (recommended)

## Installation

1. Clone the repository:
```bash
git clone git@github.com:Jamieg7/random-powerpoint.git
cd random-powerpoint
```

2. Create and activate a virtual environment:
```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install openai python-pptx requests python-dotenv
```

## API Key Configuration

You have two options for adding your OpenAI API key:

### Option 1: Using `.env` file (Recommended)

1. Create a `.env` file in the project root:
```bash
touch .env
```

2. Add your API key to the `.env` file:
```
OPENAI_API_KEY=sk-your-actual-api-key-here
```

The `.env` file is already in `.gitignore`, so your API key won't be committed to the repository.

### Option 2: Replace Hardcoded Value

If you prefer not to use environment variables, you can directly edit `generate.py` and replace the placeholder on line 18:

```python
# Change this line:
api_key = "OR_HARDCODED_OPENAI_API_KEY"

# To:
api_key = "sk-your-actual-api-key-here"
```

**Note:** If you choose this method, be careful not to commit your API key to version control!

## Usage

### Basic Command

Generate a presentation with the default category (Geography) and specify the number of slides:

```bash
python generate.py --num_slides 10
```

### With Custom Topic Category

Generate a presentation on a random topic within a specific category:

```bash
python generate.py --topic_category "Computer Science" --num_slides 15
```

### Available Slide Counts

You can generate presentations with:
- `5` slides
- `10` slides  
- `15` slides

### Example Commands

```bash
# Quick 5-slide Geography presentation
python generate.py --num_slides 5

# 10-slide presentation on a random Science topic
python generate.py --topic_category "Science" --num_slides 10

# 15-slide deep dive on a random History topic
python generate.py --topic_category "History" --num_slides 15

# Tech-focused presentation
python generate.py --topic_category "Technology" --num_slides 10
```

## Output

The script will:
1. Generate a random topic within your specified category
2. Create detailed slide content with technical depth
3. Generate AI images for each content slide (except the title slide)
4. Save the presentation as `[Topic_Name]_[num_slides]_slides.pptx` in the current directory

Example output filename: `Quantum_Entanglement_and_Bell_Theorems_10_slides.pptx`


## Project Structure

```
random-powerpoint/
├── generate.py          # Main script
├── .env                 # API key (not in git)
├── .gitignore          # Git ignore rules
├── venv/               # Virtual environment
└── README.md           # This file
```

## Notes

- The script uses `gpt-4o-mini` by default for cost-effectiveness. You can change this to `gpt-4o` in the code for better quality.
- Image generation uses DALL-E 3 (standard quality, 1024x1024)
- Generated images are temporarily downloaded and then removed after being added to the presentation
- The script creates 16:9 aspect ratio presentations (standard widescreen format)

## License

Just having fun building a tool for me, I hope it helps you! Use however you want!

## Contributing

This is a vibecoded project - contributions, improvements, and suggestions are welcome! Just open an issue or PR.

---

**Remember:** The goal is to practice presenting on topics you don't know. Embrace the challenge! 🚀


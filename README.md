Here's a GitHub description for your demo file:

---

# Excel Ollama Integration

A VBA module that connects Microsoft Excel to a local [Ollama](https://ollama.com) installation, enabling you to use large language models directly in your spreadsheets.

## Features

- **Simple LLM Integration**: Call Ollama models directly from Excel formulas
- **Intelligent Caching**: Responses are cached locally to avoid redundant API calls and improve performance
- **Range Analysis**: Pass Excel ranges directly to the model for data analysis
- **Flexible Model Selection**: Use any locally installed Ollama model (defaults to llama3)
- **Mac & Windows Compatible**: Works on both platforms with automatic path detection
- **Error Handling**: Graceful error handling with informative messages

## Functions

### `OLLAMA(prompt, [range], [model])`
Main function to interact with Ollama models.

**Parameters:**
- `prompt` (required): Your question or instruction
- `range` (optional): Excel range to include in the analysis
- `model` (optional): Model name (default: "llama3")

**Examples:**
```excel
=OLLAMA("Summarize this in 3 words", A1:A10)
=OLLAMA("What is the capital of France?")
=OLLAMA("Explain quantum computing", , "llama2")
=OLLAMA("Translate to Spanish: Hello", A1, "mistral")
```

### `OLLAMA_ANALYZE(prompt, range, [model])`
Specialized function for analyzing Excel data ranges with concise responses.

**Example:**
```excel
=OLLAMA_ANALYZE("What's the trend?", A1:A20)
```

### `ClearOllamaCache()`
VBA macro to clear the local response cache.

## Requirements

- Microsoft Excel (Mac or Windows)
- [Ollama](https://ollama.com) installed and running locally
- [VBA-Web](https://github.com/VBA-tools/VBA-Web) library for REST API calls

## Installation

1. Install Ollama from ollama.com
2. Pull your desired model (llama3 is the one this workbook uses by default): ollama pull llama3
3. Install VBA-Web library in your Excel project (https://github.com/VBA-tools/VBA-Web) and import Module1 from this workbook into your workbook, OR…
4. …just copy and edit this workbook to do what you want
5. If you access Ollama from a different machine, edit the Client.BaseUrl in the OLLAMA function in Module1
6. Start using the functions in your spreadsheets!

## How It Works

- Makes HTTP POST requests to Ollama's local API (http://localhost:11434)
- Caches responses in `~/Documents/OllamaCache.txt` to avoid repeat calls
- Only caches successful responses (errors trigger fresh API calls)
- Converts Excel ranges to CSV format for model processing

## License

The MIT License (MIT)

Copyright (c) 2025 David Griffiths

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
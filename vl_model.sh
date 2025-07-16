vllm serve Qwen/Qwen2.5-VL-7B-Instruct --port 9901 --host 0.0.0.0 --dtype bfloat16 --limit-mm-per-prompt image=5,video=5 --served-model-name CustomLLM

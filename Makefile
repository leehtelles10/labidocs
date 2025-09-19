.PHONY: run lint test docker-build docker-run

run:
	python -m streamlit run app_v2.py

lint:
	flake8 .

test:
	pytest -q

docker-build:
	docker build -t labidocs:latest .

docker-run:
	docker run --rm -p 8501:8501 labidocs:latest

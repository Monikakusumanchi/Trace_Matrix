FROM python:3.11
EXPOSE 8000
WORKDIR /code

# Copy just the requirements file and install dependencies
COPY requirements.txt ./
RUN pip install -r requirements.txt

# Copy the rest of the application files
COPY . ./
CMD ["uvicorn", "main:app", "--reload", "--host", "0.0.0.0", "--port", "8000"]
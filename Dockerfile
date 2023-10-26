# Stage 1: Build the React frontend
FROM node:12 AS ReactImage
WORKDIR /app/frontend
COPY ./frontend/package.json ./frontend/yarn.lock /app/frontend/
RUN yarn install --no-optional
COPY ./frontend ./
RUN yarn build

# Stage 2: Build the FastAPI backend
FROM python:3.8-slim-buster AS FastAPIImage
ENV PYTHONUNBUFFERED 1
ENV PYTHONDONTWRITEBYTECODE 1
RUN python3 -m pip install --upgrade pip setuptools wheel 

WORKDIR /app/api
COPY ./api/requirements.txt /app/api/
RUN pip install -r requirements.txt
COPY ./api ./

# Stage 3: Create the final image
FROM python:3.8-slim-buster
ENV PYTHONUNBUFFERED 1
ENV PYTHONDONTWRITEBYTECODE 1

WORKDIR /app/api
COPY --from=FastAPIImage /app/api /app/api
COPY --from=ReactImage /app/frontend/build /app/api/templates/build

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]

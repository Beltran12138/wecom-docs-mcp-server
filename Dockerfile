FROM python:3.11-slim
WORKDIR /app
COPY wecom_docs_mcp_server.py .
ENV NODE_EXE=node
CMD ["python", "wecom_docs_mcp_server.py"]

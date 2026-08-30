# Glama inspection / generic container build.
# The COM tool family (PDF export, slide-image export, opens-clean checks)
# requires Windows + Microsoft PowerPoint and is not available in a
# container; the file-based majority of the toolset works anywhere. The
# server starts and serves MCP over stdio.
FROM python:3.12-slim
RUN pip install --no-cache-dir kitchensink4ppt
ENTRYPOINT ["kitchensink4ppt"]

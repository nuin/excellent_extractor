from typing import List, Optional
from fastapi import FastAPI
from strawberry.fastapi import GraphQLRouter
from strawberry.types import Info
import strawberry
from excel_extractor import ExcelExtractor
from pydantic_settings import BaseSettings
from strawberry.types.info import RootValueType
from strawberry.fastapi import BaseContext
import os

class Settings(BaseSettings):
    index_directory: str = "/Users/nuin/Projects/ahs/variant_db/excellent_extractor/index/"
    base_directory: str = "/path/to/your/base/directory/"  # Update this path

    class Config:
        env_file = ".env"

settings = Settings()

@strawberry.type
class FileLocation:
    filename: str
    relative_path: str

@strawberry.type
class SearchResult:
    filename: str
    relative_path: str
    sheet_name: str
    score: float
    highlight: str

class CustomContext(BaseContext):
    def __init__(self):
        super().__init__()
        self.extractor = ExcelExtractor(settings.base_directory, settings.index_directory)

@strawberry.type
class Query:
    @strawberry.field
    def search_content(self, info: Info[CustomContext, RootValueType], query: str, limit: Optional[int] = None) -> List[SearchResult]:
        results = info.context.extractor.search(query, limit)
        return [SearchResult(**r) for r in results]

    @strawberry.field
    def get_file_location(self, info: Info[CustomContext, RootValueType], filename: str) -> Optional[FileLocation]:
        location = info.context.extractor.get_file_location(filename)
        if location:
            return FileLocation(**location)
        return None

    @strawberry.field
    def search_image_content(self, info: Info[CustomContext, RootValueType], query: str) -> List[SearchResult]:
        results = info.context.extractor.search_images(query)
        return [SearchResult(**r) for r in results]

    @strawberry.field
    def search_by_filename(self, info: Info[CustomContext, RootValueType], filename: str) -> List[FileLocation]:
        results = info.context.extractor.search_by_filename(filename)
        return [FileLocation(**r) for r in results]

    @strawberry.field
    def search_by_gene_symbol(self, info: Info[CustomContext, RootValueType], gene_symbol: str) -> List[FileLocation]:
        results = info.context.extractor.search_by_gene_symbol(gene_symbol)
        return [FileLocation(**r) for r in results]

schema = strawberry.Schema(query=Query)

async def get_context() -> CustomContext:
    return CustomContext()

app = FastAPI()

graphql_app = GraphQLRouter(
    schema,
    context_getter=get_context,
)

app.include_router(graphql_app, prefix="/graphql")

@app.get("/")
async def root():
    return {"message": "Welcome to Excel Extractor API", "index_directory": settings.index_directory}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
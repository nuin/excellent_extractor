from typing import List, Optional
from fastapi import FastAPI
from strawberry.fastapi import GraphQLRouter
import strawberry
from excel_extractor import ExcelExtractor  # Import from your Excel extractor file

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

@strawberry.type
class Query:
    @strawberry.field
    def search_content(self, query: str, limit: Optional[int] = None) -> List[SearchResult]:
        extractor = ExcelExtractor(None, 'index_directory')
        results = extractor.search(query, limit)
        return [SearchResult(**r) for r in results]

    @strawberry.field
    def get_file_location(self, filename: str) -> Optional[FileLocation]:
        extractor = ExcelExtractor(None, 'index_directory')
        location = extractor.get_file_location(filename)
        if location:
            return FileLocation(**location)
        return None

    @strawberry.field
    def search_image_content(self, query: str) -> List[SearchResult]:
        extractor = ExcelExtractor(None, 'index_directory')
        results = extractor.search_images(query)
        return [SearchResult(**r) for r in results]

schema = strawberry.Schema(query=Query)

app = FastAPI()

graphql_app = GraphQLRouter(schema)

app.include_router(graphql_app, prefix="/graphql")

@app.get("/")
async def root():
    return {"message": "Welcome to Excel Extractor API"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

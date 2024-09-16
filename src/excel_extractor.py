import os
import logging
from dataclasses import dataclass
from typing import List, Tuple

import click
import pandas as pd
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
import pytesseract
from PIL import Image
from whoosh import index
from whoosh.fields import Schema, TEXT, ID
from whoosh.qparser import QueryParser
from rich.console import Console
from rich.logging import RichHandler
from rich.progress import Progress
from rich.table import Table

@dataclass
class SheetContent:
    name: str
    cell_text: str
    images: List[Tuple[str, str]]  # (image_coord, image_text)

@dataclass
class WorkbookContent:
    filename: str
    sheets: List[SheetContent]

class ExcelExtractor:
    def __init__(self, directory_path: str, index_dir: str):
        self.directory_path = directory_path
        self.index_dir = index_dir
        self.schema = Schema(
            filename=ID(stored=True),
            sheet_name=ID(stored=True),
            content=TEXT(stored=True)
        )
        self.console = Console()

    def extract_text_from_sheet(self, sheet: pd.DataFrame) -> str:
        return sheet.to_string(index=False)

    def extract_text_from_image(self, image: Image) -> str:
        return pytesseract.image_to_string(image)

    def process_sheet(self, sheet, sheet_name: str) -> SheetContent:
        cell_text = self.extract_text_from_sheet(sheet)
        
        image_loader = SheetImageLoader(sheet)
        images = []
        for image_coord in image_loader._images:
            image = image_loader.get(image_coord)
            image_text = self.extract_text_from_image(image)
            images.append((image_coord, image_text))
        
        return SheetContent(sheet_name, cell_text, images)

    def process_workbook(self, file_path: str) -> WorkbookContent:
        filename = os.path.basename(file_path)
        sheets = []

        xlsx = pd.ExcelFile(file_path)
        for sheet_name in xlsx.sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name)
            sheet_content = self.process_sheet(xlsx.book[sheet_name], sheet_name)
            sheets.append(sheet_content)

        return WorkbookContent(filename, sheets)

    def process_directory(self) -> List[WorkbookContent]:
        workbooks = []
        excel_files = [f for f in os.listdir(self.directory_path) if f.endswith(('.xlsx', '.xls'))]
        
        with Progress() as progress:
            task = progress.add_task("[green]Processing Excel files...", total=len(excel_files))
            
            for filename in excel_files:
                file_path = os.path.join(self.directory_path, filename)
                logging.info(f"Processing {filename}...")
                workbook_content = self.process_workbook(file_path)
                workbooks.append(workbook_content)
                progress.update(task, advance=1)
        
        return workbooks

    def index_content(self, workbooks: List[WorkbookContent]):
        if not os.path.exists(self.index_dir):
            os.mkdir(self.index_dir)
        ix = index.create_in(self.index_dir, self.schema)

        with Progress() as progress:
            task = progress.add_task("[green]Indexing content...", total=sum(len(wb.sheets) for wb in workbooks))
            
            writer = ix.writer()
            for workbook in workbooks:
                for sheet in workbook.sheets:
                    content = f"{sheet.cell_text}\n"
                    for _, image_text in sheet.images:
                        content += f"{image_text}\n"
                    writer.add_document(
                        filename=workbook.filename,
                        sheet_name=sheet.name,
                        content=content
                    )
                    progress.update(task, advance=1)
            writer.commit()

    def search(self, query_str: str, limit: int = 10):
        ix = index.open_dir(self.index_dir)
        with ix.searcher() as searcher:
            query = QueryParser("content", ix.schema).parse(query_str)
            results = searcher.search(query, limit=limit)
            
            table = Table(title=f"Search Results for '{query_str}'")
            table.add_column("File", style="cyan")
            table.add_column("Sheet", style="magenta")
            table.add_column("Score", justify="right", style="green")
            table.add_column("Highlights", style="yellow")
            
            for result in results:
                table.add_row(
                    result['filename'],
                    result['sheet_name'],
                    f"{result.score:.2f}",
                    result.highlights("content")
                )
            
            self.console.print(table)

@click.group()
@click.option('--directory', type=click.Path(exists=True), help='Directory containing Excel files')
@click.option('--index-dir', type=click.Path(), default='index_directory', help='Directory to store the search index')
@click.option('--log-level', type=click.Choice(['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']), default='INFO', help='Set the logging level')
@click.pass_context
def cli(ctx, directory, index_dir, log_level):
    ctx.ensure_object(dict)
    ctx.obj['directory'] = directory
    ctx.obj['index_dir'] = index_dir
    
    logging.basicConfig(
        level=log_level,
        format="%(message)s",
        datefmt="[%X]",
        handlers=[RichHandler(rich_tracebacks=True)]
    )

@cli.command()
@click.pass_context
def process(ctx):
    """Process Excel files and create search index"""
    extractor = ExcelExtractor(ctx.obj['directory'], ctx.obj['index_dir'])
    workbooks = extractor.process_directory()
    extractor.index_content(workbooks)
    click.echo("Processing and indexing completed.")

@cli.command()
@click.argument('query')
@click.option('--limit', default=10, help='Maximum number of results to display')
@click.pass_context
def search(ctx, query, limit):
    """Search indexed Excel content"""
    extractor = ExcelExtractor(ctx.obj['directory'], ctx.obj['index_dir'])
    extractor.search(query, limit)

if __name__ == "__main__":
    cli(obj={})

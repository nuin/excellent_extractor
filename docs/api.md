# Excel Extractor API Search Documentation for Genetic Data

This document outlines all possible searches available through the Excel Extractor API, tailored for genetic data and gene variants. The API is implemented using GraphQL and provides three main query operations.

## GraphQL Endpoint

The GraphQL endpoint is available at:

```
/graphql
```

## Available Queries

### 1. Search Content

This query allows you to search for genetic content within Excel files, such as specific genes, variants, or genetic markers.

```graphql
query SearchContent($query: String!, $limit: Int) {
  searchContent(query: $query, limit: $limit) {
    filename
    relativePath
    sheetName
    score
    highlight
  }
}
```

- **Parameters**:
  - `query` (String, required): The search term or phrase (e.g., gene name, variant ID)
  - `limit` (Int, optional): Maximum number of results to return

- **Returns**: A list of `SearchResult` objects containing:
  - `filename`: Name of the Excel file containing genetic data
  - `relativePath`: Relative path to the file
  - `sheetName`: Name of the sheet where the genetic information was found
  - `score`: Relevance score of the search result
  - `highlight`: Highlighted excerpt of the matching genetic content

### 2. Get File Location

This query retrieves the location information for a specific Excel file containing genetic data.

```graphql
query GetFileLocation($filename: String!) {
  getFileLocation(filename: $filename) {
    filename
    relativePath
  }
}
```

- **Parameters**:
  - `filename` (String, required): Name of the Excel file to locate

- **Returns**: A `FileLocation` object containing:
  - `filename`: Name of the Excel file
  - `relativePath`: Relative path to the file

### 3. Search Image Content

This query allows you to search for content within images embedded in Excel files, which might include genetic diagrams, karyotypes, or other visual genetic data.

```graphql
query SearchImageContent($query: String!) {
  searchImageContent(query: $query) {
    filename
    relativePath
    sheetName
    score
    highlight
  }
}
```

- **Parameters**:
  - `query` (String, required): The search term or phrase for image content related to genetic data

- **Returns**: A list of `SearchResult` objects containing:
  - `filename`: Name of the Excel file
  - `relativePath`: Relative path to the file
  - `sheetName`: Name of the sheet where the genetic image was found
  - `score`: Relevance score of the search result
  - `highlight`: Highlighted excerpt of the matching image content description

## Example Usage

Here are some example GraphQL queries you can use:

1. Search for content related to a specific gene variant:

```graphql
query {
  searchContent(query: "BRCA1 c.5266dupC", limit: 5) {
    filename
    relativePath
    sheetName
    score
    highlight
  }
}
```

2. Get file location for a genetic study dataset:

```graphql
query {
  getFileLocation(filename: "breast_cancer_variants_2023.xlsx") {
    filename
    relativePath
  }
}
```

3. Search image content for karyotype diagrams:

```graphql
query {
  searchImageContent(query: "karyotype Down syndrome") {
    filename
    relativePath
    sheetName
    score
    highlight
  }
}
```

Remember to replace the example values with your actual search terms or filenames when using these queries. These queries can help you efficiently search through Excel files containing genetic data, locate specific files, and find relevant images or diagrams related to genetic information.
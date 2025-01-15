import asyncio
import aiohttp
import json
import os
import re
import requests

from docx2pdf import convert
from termcolor import colored


graphql_endpoint = "https://create.microsoft.com/api/graphql"
request_body = [
    {
        "operationName": "getSearchTemplateGrid",
        "variables": {
            "query": "resumes,resume,cv",
            "filters": [
                "keywords=resumes",
                "keywords=resume",
                "keywords=cv"
            ],
            "locale": "en-us",
            "offset": 0,
            "limit": 50,
            "generic": False
        },
        "query": "query getSearchTemplateGrid($query: String!, $filters: [String!], $offset: Int, $limit: Int, $locale: String, $generic: Boolean, $collectionId: String, $orderSeed: String) {\n  searchTemplates(\n    query: $query\n    filters: $filters\n    locale: $locale\n    offset: $offset\n    limit: $limit\n    generic: $generic\n    collectionId: $collectionId\n    orderSeed: $orderSeed\n  ) {\n    id\n    ...SearchTemplateGrid_searchTemplates\n    __typename\n  }\n  componentContent {\n    id\n    ...SearchTemplateGrid_componentContent\n    __typename\n  }\n}\n\nfragment SearchTemplateGrid_searchTemplates on SearchTemplates {\n  templates {\n    templates {\n      id\n      ...TemplateThumbnailCard_template\n      __typename\n    }\n    __typename\n  }\n  totalCount\n  searchStatus\n  __typename\n}\n\nfragment SearchTemplateGrid_componentContent on ComponentContent {\n  ...TemplateThumbnailCard_componentContent\n  ...noResultsHeading_ComponentContent\n  __typename\n}\n\nfragment TemplateThumbnailCard_template on Template {\n  title\n  longFormTitle\n  premium\n  templateContentType\n  ...TemplateThumbnail_template\n  ...TemplateThumbnailActions_template\n  __typename\n}\n\nfragment TemplateThumbnail_template on Template {\n  thumbnails {\n    alt\n    height\n    size\n    uri\n    width\n    contentType\n    __typename\n  }\n  __typename\n}\n\nfragment TemplateThumbnailActions_template on Template {\n  ...TemplateThumbnailContentTypeLabel_template\n  __typename\n}\n\nfragment TemplateThumbnailContentTypeLabel_template on Template {\n  templateContentType\n  supportingApplication\n  __typename\n}\n\nfragment TemplateThumbnailCard_componentContent on ComponentContent {\n  ...TemplateThumbnailActions_componentContent\n  __typename\n}\n\nfragment TemplateThumbnailActions_componentContent on ComponentContent {\n  ...TemplateThumbnailContentTypeLabel_componentContent\n  __typename\n}\n\nfragment TemplateThumbnailContentTypeLabel_componentContent on ComponentContent {\n  templateThumbnailOverlay {\n    ctaLabelPrefix\n    __typename\n  }\n  __typename\n}\n\nfragment noResultsHeading_ComponentContent on ComponentContent {\n  noSearchResults {\n    heading\n    subheading\n    __typename\n  }\n  __typename\n}"
    }
]
template_download_dir = "templates"
template_pdf_dir = "templates_pdf"
offsets = [0, 50, 100]
wdFormatPDF = 17


def format_temlate_title(template_title: str) -> str:
    return re.sub(r"\s+", "-", template_title.lower())


def get_resume_results_ids_from_graphql(offset: int) -> list[object]:
    request_body[0]["variables"]["offset"] = offset

    graphql_request = requests.post(
        graphql_endpoint,
        json=request_body
    )
    graphql_request.raise_for_status()

    graphql_response = json.loads(graphql_request.text)

    return [{"id": template["id"], "title": format_temlate_title(template["title"])} for template in filter(lambda template: template["supportingApplication"] == "WORD", graphql_response[0]["data"]["searchTemplates"]["templates"]["templates"])]


async def download_template_to_download_folder(template_id: str, formatted_template_title: str) -> None:
    if os.path.exists(template_download_dir):
        async with aiohttp.ClientSession() as session:
            async with session.get(f"https://create.microsoft.com/en-us/template/{formatted_template_title}-{template_id}") as response:
                if response.status == 200:
                    response_text = await response.text()
                    template_download_link = re.search(
                        r""",{"__typename":"TemplateAffordance","link":"(.*)","type":"DOWNLOAD"}]""", response_text).group(1)

                    async with session.get(template_download_link) as download_response:
                        if download_response.status == 200:
                            download_response = await download_response.read()
                            template_file_name = f"{
                                formatted_template_title}.docx"

                            with open(os.path.join(template_download_dir, template_file_name), "wb") as template_file:
                                template_file.write(download_response)
    else:
        print(
            colored(f"Folder {template_download_dir} does not exist!", "red"))
        print(colored("Creating folder...", "cyan"))
        os.mkdir(template_download_dir)

        await download_template_to_download_folder(template_id, formatted_template_title)


def convert_docx_to_pdf() -> None:
    if os.path.exists(template_pdf_dir):
        for template_file in os.listdir(template_download_dir):
            if template_file.endswith(".docx"):
                convert(os.path.join(template_download_dir, template_file), os.path.join(
                        template_pdf_dir, template_file.replace(".docx", ".pdf")))
    else:
        print(colored(f"Folder {template_pdf_dir} does not exist!", "red"))
        print(colored("Creating folder...", "cyan"))
        os.mkdir(template_pdf_dir)

        convert_docx_to_pdf()


async def main() -> None:
    print(colored("Starting scraper...", "cyan"))
    for page_offset in offsets:
        resume_ids = get_resume_results_ids_from_graphql(page_offset)
        print(colored(f"Downloading {len(resume_ids)} from page {
            offsets.index(page_offset) + 1} resumes...", "cyan"))
        download_pool = [download_template_to_download_folder(
            resume_id["id"], resume_id["title"]) for resume_id in resume_ids]
        await asyncio.gather(*download_pool)
        print(colored(f"Finished downloading {len(resume_ids)} resumes from page {
            offsets.index(page_offset) + 1} resumes...", "cyan"))

    print(colored("Converting docx to pdf...", "cyan"))
    convert_docx_to_pdf()
    print(colored("Done!", "green"))


if __name__ == "__main__":
    asyncio.run(main())

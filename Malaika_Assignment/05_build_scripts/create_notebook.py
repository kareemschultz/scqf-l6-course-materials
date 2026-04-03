import asyncio, os
from pathlib import Path
from notebooklm import NotebookLMClient

PDF_DIR = Path(__file__).parent / 'course_pdfs'
PDFS = sorted(PDF_DIR.glob('*.pdf'))

YOUTUBE_LINKS = [
    ('Topic1 - Evolution of HRM',                    'https://www.youtube.com/watch?v=Kxc8KceOb14'),
    ('Topic1 - Transformation Personnel to HRM',     'https://www.youtube.com/watch?v=8ReX2poQyJ0'),
    ('Topic2 - HR Strategy and Planning',            'https://www.youtube.com/watch?v=8mwCiDKgNd4'),
    ('Topic3 - HRP Programme Implementation',        'https://www.youtube.com/watch?v=ha2ZCiWKtTU'),
    ('Topic4 - Job Analysis',                        'https://www.youtube.com/watch?v=oas5n1nFHQQ'),
    ('Topic4 - Job Design',                          'https://www.youtube.com/watch?v=uUG-Z5sg2UM'),
]

async def main():
    async with await NotebookLMClient.from_storage() as client:
        print('Creating notebook: MGMT268 - Malaika Assignment')
        nb = await client.notebooks.create('MGMT268 - Behavioral vs Rational Job Design')
        print(f'Notebook created: {nb.id}')

        # Add PDFs
        for pdf in PDFS:
            print(f'Adding PDF: {pdf.name}...')
            try:
                await client.sources.add_file(nb.id, str(pdf))
                print(f'  Added: {pdf.name}')
            except Exception as e:
                print(f'  Error: {e}')

        # Add YouTube videos
        for title, url in YOUTUBE_LINKS:
            print(f'Adding YouTube: {title}...')
            try:
                await client.sources.add_youtube(nb.id, url)
                print(f'  Added: {title}')
            except Exception as e:
                print(f'  Trying add_url for {title}: {e}')
                try:
                    await client.sources.add_url(nb.id, url)
                    print(f'  Added via URL: {title}')
                except Exception as e2:
                    print(f'  Failed: {e2}')

        print(f'\nDone! Notebook ID: {nb.id}')
        print('Open at: https://notebooklm.google.com')

asyncio.run(main())

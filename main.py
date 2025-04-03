from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse
from fastapi import Request
from pathlib import Path
import win32com.client
import pythoncom
import tempfile
import shutil
import os
import zipfile
import logging
from typing import List
import sys

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('converter.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = FastAPI()
templates = Jinja2Templates(directory="templates")

# Configuração de diretórios
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Servir arquivos estáticos
app.mount("/output", StaticFiles(directory="output"), name="output")

def setup_directories():
    """Garante que os diretórios necessários existam"""
    try:
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        logger.info("Diretórios configurados com sucesso")
    except Exception as e:
        logger.error(f"Erro ao criar diretórios: {str(e)}")
        raise

def validate_file(file: UploadFile):
    """Valida o arquivo enviado"""
    if not file.filename.lower().endswith(".pptx"):
        logger.error(f"Tipo de arquivo inválido: {file.filename}")
        raise HTTPException(status_code=400, detail="Apenas arquivos .pptx são suportados")
    
    # Verifica tamanho máximo (10MB)
    max_size = 10 * 1024 * 1024
    file.file.seek(0, os.SEEK_END)
    file_size = file.file.tell()
    file.file.seek(0)
    
    if file_size > max_size:
        logger.error(f"Arquivo muito grande: {file_size} bytes")
        raise HTTPException(status_code=400, detail="Tamanho máximo do arquivo é 10MB")

def convert_with_powerpoint(input_path: Path, output_dir: Path, format: str = "png") -> List[Path]:
    """Converte PPTX para imagens usando Microsoft PowerPoint com tratamento robusto de erros"""
    powerpoint = None
    presentation = None
    try:
        pythoncom.CoInitialize()
        logger.info(f"Iniciando conversão do arquivo: {input_path}")
        
        # Verifica se o arquivo existe
        if not input_path.exists():
            raise FileNotFoundError(f"Arquivo {input_path} não encontrado")

        # Cria diretório de saída com permissões adequadas
        output_dir.mkdir(exist_ok=True, parents=True)
        logger.info(f"Diretório de saída: {output_dir.absolute()}")

        # Verifica permissões de escrita
        test_file = output_dir / "test_permission.tmp"
        try:
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
        except Exception as e:
            raise RuntimeError(f"Sem permissão para escrever no diretório de saída: {str(e)}")

        # Inicializa PowerPoint
        try:
            logger.info("Iniciando PowerPoint...")
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            
            # Tenta configurar como invisível (ignora erro se não for possível)
            try:
                powerpoint.Visible = 0
            except:
                logger.warning("PowerPoint não pode ser executado em modo invisível")
                powerpoint.Visible = 1
            
            logger.info(f"PowerPoint versão: {powerpoint.Version}")
        except Exception as e:
            raise RuntimeError(f"Falha ao iniciar PowerPoint: {str(e)}")

        output_paths = []
        
        try:
            logger.info(f"Abrindo apresentação: {input_path}")
            presentation = powerpoint.Presentations.Open(str(input_path))
            total_slides = presentation.Slides.Count
            logger.info(f"Total de slides: {total_slides}")

            # Exportar cada slide
            for i in range(1, total_slides + 1):
                output_filename = f"slide_{i}.{format}"
                output_path = output_dir / output_filename
                logger.info(f"Convertendo slide {i} -> {output_path}")

                try:
                    # Caminho absoluto para evitar problemas de permissão
                    abs_output_path = output_path.absolute()
                    
                    # Remove arquivo existente se houver
                    if abs_output_path.exists():
                        os.remove(abs_output_path)

                    # Exporta o slide
                    presentation.Slides(i).Export(
                        str(abs_output_path),
                        format.upper(),  # PNG ou JPG
                        1920,  # Width
                        1080   # Height
                    )

                    # Verifica se o arquivo foi criado
                    if not abs_output_path.exists():
                        logger.error(f"Arquivo de saída não foi criado: {abs_output_path}")
                        continue

                    output_paths.append(abs_output_path)
                    logger.info(f"Slide {i} convertido com sucesso")

                except Exception as slide_error:
                    logger.error(f"Erro no slide {i}: {str(slide_error)}", exc_info=True)
                    continue

            if not output_paths:
                raise RuntimeError("Nenhum slide foi convertido com sucesso")

            return output_paths

        except Exception as e:
            logger.error(f"Erro durante a conversão: {str(e)}", exc_info=True)
            raise
        finally:
            if presentation:
                try:
                    logger.info("Fechando apresentação...")
                    presentation.Close()
                except Exception as e:
                    logger.error(f"Erro ao fechar apresentação: {str(e)}")
    except Exception as e:
        logger.error(f"Erro na conversão: {str(e)}", exc_info=True)
        raise
    finally:
        if powerpoint:
            try:
                logger.info("Fechando PowerPoint...")
                powerpoint.Quit()
            except Exception as e:
                logger.error(f"Erro ao fechar PowerPoint: {str(e)}")
        pythoncom.CoUninitialize()

def create_zip(file_paths: List[Path], zip_path: Path):
    """Cria um arquivo ZIP com os arquivos convertidos"""
    try:
        logger.info(f"Criando arquivo ZIP: {zip_path}")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in file_paths:
                zipf.write(file, arcname=file.name)
        logger.info(f"ZIP criado com {len(file_paths)} arquivos")
    except Exception as e:
        logger.error(f"Erro ao criar ZIP: {str(e)}")
        raise

@app.post("/upload/")
async def upload_pptx(file: UploadFile = File(...), format: str = "png"):
    """Endpoint para upload e conversão de arquivos PPTX"""
    logger.info(f"Iniciando upload do arquivo: {file.filename}")
    
    try:
        # Validações iniciais
        validate_file(file)
        setup_directories()

        # Limpar diretório de saída
        shutil.rmtree(OUTPUT_DIR, ignore_errors=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        # Salvar arquivo temporariamente
        temp_dir = tempfile.mkdtemp()
        try:
            filepath = Path(temp_dir) / file.filename
            logger.info(f"Salvando arquivo temporário em: {filepath}")
            
            with filepath.open("wb") as buffer:
                shutil.copyfileobj(file.file, buffer)

            # Converter usando PowerPoint
            logger.info(f"Iniciando conversão para formato: {format}")
            output_files = convert_with_powerpoint(filepath, Path(OUTPUT_DIR), format)
            
            if not output_files:
                raise HTTPException(status_code=500, detail="Nenhum slide foi convertido")

            # Criar ZIP com resultados
            zip_filename = f"{filepath.stem}_slides.zip"
            zip_path = Path(OUTPUT_DIR) / zip_filename
            create_zip(output_files, zip_path)

            # Preparar URLs para resposta
            image_urls = [f"/output/{f.name}" for f in output_files]
            zip_url = f"/output/{zip_filename}"

            logger.info(f"Conversão concluída. Slides gerados: {len(output_files)}")
            
            return JSONResponse(content={
                "images": image_urls,
                "zip_url": zip_url,
                "message": f"Convertido {len(output_files)} slides com sucesso!"
            })

        except Exception as e:
            logger.error(f"Erro durante o processamento: {str(e)}", exc_info=True)
            raise HTTPException(status_code=500, detail=f"Erro na conversão: {str(e)}")
        finally:
            # Limpar arquivos temporários
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro inesperado: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail="Erro interno no servidor")

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    """Endpoint principal que serve a página HTML"""
    return templates.TemplateResponse("index.html", {"request": request})

if __name__ == "__main__":
    try:
        logger.info("Iniciando servidor FastAPI...")
        import uvicorn
        uvicorn.run(app, host="0.0.0.0", port=5000, log_level="info")
    except Exception as e:
        logger.error(f"Falha ao iniciar servidor: {str(e)}", exc_info=True)
        sys.exit(1)
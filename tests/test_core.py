import os
import tempfile
import pytest

from .. import app


def test_importa_app():
    """Confere se o app importa sem erros."""
    assert hasattr(app, "extrair_metadados")
    assert hasattr(app, "processar_projeto")


def test_extrair_metadados_vazio():
    """Chama extrair_metadados em um arquivo inexistente e espera lista vazia."""
    resultado = app.extrair_metadados("arquivo_que_nao_existe.tmdl")
    assert resultado == ([], [], [], [])
    # assert isinstance(resultado, list)


def test_processar_projeto_zip_vazio():
    """Cria um zip vazio temporário e testa processar_projeto."""
    with tempfile.NamedTemporaryFile(suffix=".zip") as tmpzip:
        resultado = app.processar_projeto(tmpzip.name, [], None)
        assert resultado is None

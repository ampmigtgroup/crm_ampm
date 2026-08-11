#!/usr/bin/env python3

from __future__ import annotations

import ast
import importlib.util
import py_compile
import sys
from pathlib import Path


REQUIRED_SYMBOLS = {
    "CAMINHO_ARQUIVO",
    "COLUNAS_FILA",
    "ENTIDADES",
    "exigir_login",
    "inicializar_estado",
    "construir_base_unificada",
    "salvar_fila_no_disco",
}


def fail(message: str) -> None:
    print(f"[FALHA] {message}")
    raise SystemExit(1)


def main() -> int:

    app = Path(
        sys.argv[1]
        if len(sys.argv) > 1
        else "app_crm.py"
    )

    if not app.exists():
        fail(f"arquivo não encontrado: {app}")

    # 1. Teste de compilação Python
    try:
        py_compile.compile(
            str(app),
            doraise=True
        )
    except py_compile.PyCompileError as exc:
        fail(
            f"erro de sintaxe/compilação:\n{exc}"
        )

    # 2. Analisa a estrutura do código
    try:
        source = app.read_text(
            encoding="utf-8"
        )

        tree = ast.parse(
            source,
            filename=str(app)
        )

    except SyntaxError as exc:
        fail(
            f"AST inválida: "
            f"linha {exc.lineno}: {exc.msg}"
        )

    # 3. Descobre símbolos definidos
    defined = set()

    for node in tree.body:

        if isinstance(
            node,
            (
                ast.FunctionDef,
                ast.AsyncFunctionDef,
                ast.ClassDef,
            ),
        ):
            defined.add(node.name)

        elif isinstance(node, ast.Assign):

            for target in node.targets:

                if isinstance(
                    target,
                    ast.Name,
                ):
                    defined.add(target.id)

        elif isinstance(
            node,
            ast.AnnAssign,
        ):

            if isinstance(
                node.target,
                ast.Name,
            ):
                defined.add(
                    node.target.id
                )

    # 4. Verifica componentes críticos
    missing = sorted(
        REQUIRED_SYMBOLS - defined
    )

    if missing:

        fail(
            "símbolos essenciais ausentes: "
            + ", ".join(missing)
        )

    # 5. Verifica dependências principais
    imports = []

    for node in tree.body:

        if isinstance(
            node,
            ast.Import,
        ):

            imports.extend(
                alias.name.split(".")[0]
                for alias in node.names
            )

        elif isinstance(
            node,
            ast.ImportFrom,
        ):

            if node.module:

                imports.append(
                    node.module.split(".")[0]
                )

    dependencies = {
        "streamlit",
        "pandas",
        "pydeck",
        "requests",
        "openpyxl",
        "streamlit_authenticator",
    }

    missing_packages = sorted(
        name
        for name in set(imports)
        if name in dependencies
        and importlib.util.find_spec(name) is None
    )

    if missing_packages:

        print(
            "[AVISO] Dependências não instaladas:"
        )

        for package in missing_packages:
            print(f"  - {package}")

    print()
    print("=" * 60)
    print("       VALIDAÇÃO DO CRM AMPM")
    print("=" * 60)
    print("[OK] compilação Python")
    print("[OK] análise AST")
    print(
        f"[OK] símbolos essenciais: "
        f"{len(REQUIRED_SYMBOLS)}"
    )
    print(
        f"[OK] arquivo validado: {app}"
    )
    print("=" * 60)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

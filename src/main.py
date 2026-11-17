#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Keno NC – mobil (Android) változat (Norbi & Claude)

A számolás a keno_core.py-ban van (ugyanaz, mint az asztali programban).
Ez a fájl csak a gombos felület.

Adatok a telefonon:  /Norbi/Keno/   (keno.db, keno.csv, Keno.xlsx)
"""

import asyncio
import os

import flet as ft

import keno_core as core

MOBIL_VERZIO = "1.0"

ANDROID_MAPPA = "/storage/emulated/0/Norbi/Keno"


def _irhato(mappa):
    """Igaz, ha a mappa létrehozható és írható."""
    try:
        os.makedirs(mappa, exist_ok=True)
        proba = os.path.join(mappa, ".iras_proba")
        with open(proba, "w") as f:
            f.write("ok")
        os.remove(proba)
        return True
    except Exception:
        return False


async def main(page: ft.Page):
    page.title = "Keno NC"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 16
    page.scroll = ft.ScrollMode.AUTO

    android = page.platform in (ft.PagePlatform.ANDROID, ft.PagePlatform.ANDROID_TV)
    if android:
        adat_mappa = ANDROID_MAPPA
    else:
        # Gépen próbafuttatáshoz: a program melletti "keno_adat" mappa
        adat_mappa = os.path.join(os.path.dirname(os.path.abspath(__file__)), "keno_adat")

    # ── Engedély (Minden fájl kezelése) ─────────────────────────────────────
    ph = None
    if android:
        try:
            import flet_permission_handler as fph
            ph = fph.PermissionHandler()
        except Exception:
            ph = None

    # ── Vezérlők ────────────────────────────────────────────────────────────
    cim = ft.Text(f"Keno NC  mobil v{MOBIL_VERZIO}", size=24, weight=ft.FontWeight.BOLD)
    alcim = ft.Text(f"core v{core.VERZIO}  ·  {adat_mappa}", size=12, color=ft.Colors.GREY_700)

    lbl_huzasok = ft.Text("", size=16)
    lbl_utolso  = ft.Text("", size=16)
    lbl_export  = ft.Text("", size=16)

    tf_napok = ft.TextField(
        label="Napok (utolsó N húzás)",
        width=170,
        keyboard_type=ft.KeyboardType.NUMBER,
        input_filter=ft.NumbersOnlyInputFilter(),
        text_align=ft.TextAlign.CENTER,
    )
    btn_napok   = ft.FilledButton("Alkalmaz")
    btn_sorrend = ft.OutlinedButton("", icon=ft.Icons.SWAP_VERT)
    btn_letolt  = ft.FilledButton("Letöltés", icon=ft.Icons.DOWNLOAD, height=60, expand=True)
    btn_excel   = ft.FilledButton("Excel készítése", icon=ft.Icons.TABLE_CHART, height=60, expand=True)
    toltes      = ft.ProgressRing(visible=False, width=28, height=28)
    lbl_allapot = ft.Text("", size=15, selectable=True)

    engedely_doboz = ft.Container(visible=False)

    # ── Segédek ─────────────────────────────────────────────────────────────
    def uzenet(szoveg, hiba=False):
        page.show_dialog(ft.SnackBar(
            ft.Text(szoveg),
            bgcolor=ft.Colors.RED_700 if hiba else None,
            show_close_icon=True,
        ))

    def allapot_frissit():
        try:
            total, utolso = core.db_info()
        except Exception:
            total, utolso = 0, ""
        lbl_huzasok.value = f"Húzások az adatbázisban: {total}"
        lbl_utolso.value  = f"Utolsó húzás: {utolso or '–'}"
        lbl_export.value  = f"Utolsó Excel: {core.cfg_get('utolso_export', '–')}"
        tf_napok.value    = str(core.napok())
        btn_sorrend.content = ("Legújabb felül" if core.legujabb_felul()
                               else "Legrégebbi felül")
        page.update()

    def gombok(engedve):
        for b in (btn_letolt, btn_excel, btn_napok, btn_sorrend):
            b.disabled = not engedve
        toltes.visible = not engedve
        page.update()

    # ── Gombok ──────────────────────────────────────────────────────────────
    def napok_alkalmaz(_=None):
        try:
            n = int(tf_napok.value or "0")
        except ValueError:
            n = 0
        if n < 1:
            uzenet("A napok száma legalább 1 legyen.", hiba=True)
            tf_napok.value = str(core.napok())
            page.update()
            return
        core.cfg_set("napok", n)
        uzenet(f"Napok: {n}")
        allapot_frissit()

    def sorrend_valt(_=None):
        core.cfg_set("sorrend_desc", "0" if core.legujabb_felul() else "1")
        allapot_frissit()

    async def letoltes(_=None):
        gombok(False)
        lbl_allapot.value = "Letöltés folyamatban..."
        page.update()
        try:
            meret, new, skip = await asyncio.to_thread(core.letoltes)
            lbl_allapot.value = (f"Letöltve: {meret / 1024:.1f} KB\n"
                                 f"Új húzások: {new}\n"
                                 f"Már létezett: {skip}")
        except Exception as e:
            lbl_allapot.value = f"Letöltési hiba:\n{e}"
            uzenet("Nem sikerült a letöltés.", hiba=True)
        gombok(True)
        allapot_frissit()

    async def excel(_=None):
        gombok(False)
        lbl_allapot.value = "Excel készül..."
        page.update()
        ok, szoveg = await asyncio.to_thread(core.excel_keszites)
        lbl_allapot.value = szoveg
        if not ok:
            uzenet("Az Excel nem készült el.", hiba=True)
        gombok(True)
        allapot_frissit()

    tf_napok.on_submit   = napok_alkalmaz
    btn_napok.on_click   = napok_alkalmaz
    btn_sorrend.on_click = sorrend_valt
    btn_letolt.on_click  = letoltes
    btn_excel.on_click   = excel

    fo_resz = ft.Column(
        visible=False,
        spacing=14,
        controls=[
            ft.Container(
                padding=12, border_radius=10, bgcolor=ft.Colors.BLUE_50,
                content=ft.Column([lbl_huzasok, lbl_utolso, lbl_export], spacing=4),
            ),
            ft.Row([tf_napok, btn_napok], vertical_alignment=ft.CrossAxisAlignment.CENTER),
            btn_sorrend,
            ft.Row([btn_letolt]),
            ft.Row([btn_excel]),
            ft.Row([toltes, ft.Container(lbl_allapot, expand=True)]),
        ],
    )

    # ── Indítás: engedély ellenőrzése, adatbázis előkészítése ───────────────
    def indulas_kesz():
        core.init(adat_mappa)
        engedely_doboz.visible = False
        fo_resz.visible = True
        allapot_frissit()

    async def engedely_kerese(_=None):
        if ph is not None:
            try:
                import flet_permission_handler as fph
                await ph.request(fph.Permission.MANAGE_EXTERNAL_STORAGE)
            except Exception:
                try:
                    await ph.open_app_settings()
                except Exception:
                    pass
        await asyncio.sleep(0.5)
        if _irhato(adat_mappa):
            indulas_kesz()
        else:
            uzenet("Még nincs engedély. Kapcsold be, majd nyomd meg újra.", hiba=True)

    engedely_doboz.content = ft.Column(spacing=12, controls=[
        ft.Text("Engedély szükséges", size=20, weight=ft.FontWeight.BOLD),
        ft.Text("Az app a /Norbi/Keno mappába menti a keno.db-t és a Keno.xlsx-et.\n"
                "Ehhez kapcsold be: Minden fájl kezelése (All files access).\n\n"
                "Ha a gomb nem nyitja meg a beállítást: Beállítások → Alkalmazások → "
                "Keno NC → Engedélyek → Fájlok → Minden fájl kezelése.",
                size=15),
        ft.FilledButton("Engedély megadása", height=52, on_click=engedely_kerese),
    ])

    page.add(cim, alcim, ft.Divider(), engedely_doboz, fo_resz)

    if _irhato(adat_mappa):
        try:
            indulas_kesz()
        except Exception as e:
            page.add(ft.Text(f"Hiba induláskor:\n{e}", color=ft.Colors.RED_700, selectable=True))
            page.update()
    else:
        engedely_doboz.visible = True
        page.update()


if __name__ == "__main__":
    ft.run(main)

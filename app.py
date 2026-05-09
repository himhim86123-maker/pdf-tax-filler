import streamlit as st
import fitz
import io
import zipfile
import re
import os
import base64
import tempfile

st.set_page_config(page_title="PDF智能填表系统 v14.2", layout="wide")

# ---- 字符宽度常量 ----
CHAR_W = 4.0
COMMA_W = 2.5
DOT_W = 2.0
FONTSIZE = 8.0

# ---- 内嵌逗號子集字體（11.1KB base64，零外部依賴） ----
COMMA_FONT_B64 = (
    "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"
    "/8BAKQwATQMpCSUcFiMkaSEjaiAhBh8cBhZJbAEGDCFzAQYcAEohAXQhBAYNAD8/Kysr"
    "ETkREjkrKxESOQEvEM4rETk5Ly8Q7SsRMxDtMDErKysrKysAKzcyFhUUBiMiJjU0NjMy"
    "FhUUBwYVFBYzMjY1NCYjIgYHJzczByMHNjZFExgZGhIcBwYGBAMDEwsQEQ8PCxAHCQVS"
    "A0gEChNyHRgbIBUPBgkIAwYDBAQICBoYFBkIDQFbD0IOBgAAAgAMAAIAcwCxAAsALQDf"
    "uQAp//CzCwBNGrj/6LQLEgBMF7j/8EAPCwBNExAWAE0TEA4QAEwRuP/wtBESAEwNuP/o"
    "sxcATQ24//C0ERIATAq4//BAERQVAEwIEA4QAEwEEA4QAEwCuP/otBQVAEwCuP/wtBET"
    "AEwCuP/wQBANAE0hHRsbFQYdDwArHRUVuP/Asw4ATRW4/8CzDABND7j/wEAiFQBNDy8V"
    "KwkDGChJbAESHhhzARIJDEoYAXQYBRIDSWwSDQA/Kz8rKysREjkBLxDOKysrEO0yEO0S"
    "OS/tMDErKysrKysrKysrKysrKzcUFjMyNjU0JiMiBjcyFhUUBiMiJjU0NjMyFhUUBiMi"
    "JjU0NzQ1NCYjIgYVNjYfFQ8PEBEMCxMfGBUdExscIh0OEQUGBQUBBggRGAYVSB4hFh4U"
    "Fg0XIBQeHSknLjEQBgYGBQUCAgICBAUfNA0NAAABABAAAgBxAK8AFwBItxYIFwBNEAAH"
    "uP/AQCQXAE0HHQ1AFwBNDUATFABMDQ0VABkVFBVpERRqCg0WEUtsFgQAPys/KysBLxDO"
    "ETkvKyvtKxEzMDErNwYGFRQXFBUUBiMiJjU0NjcjIgYHJzczcRUZAQUGBQcRJTILDQMG"
    "B1qoIkEYBAcICggGBwkTOEMODgEqAAMACwACAHEAsQALACMALwGmtS8wFwBNLLj/8EAU"
    "FgBNIxAUFQBMIBAXAE0WCBcATQm4//BAIBAATQgoFwBNCCgUFQBMCCAWAE0IIBETAEwH"
    "OBQWAEwsuP/AtBUWAEwsuP/AsxMATSy4/+CzFwBNLLj/4LMUAE0suP/gtA0QAEwouP/o"
    "sxcATSi4/+hADhAATSYYFwBNJhgQAE0iuP/4QA4QAE0iEAwATRQYEABNHbj/6EALERIA"
    "TBkYERIATBC4//BAChcATRAYERIATA64/+hAGRESAEwKGBAATQggDABNCBALAE0EGBAA"
    "TQK4/+CzFwBNArj/6LMQAE0CuP/wQBoODwBMIRUMEiQdHgYdDCodGAAdEh4YHhgSErj/"
    "wLMOAE0SuP/AQAkMAE0SQAsATQy4/8C2FQBNDDESLbj/8LQWFwBMLbj/8LMUAE0tuP/o"
    "QBkTAE0JGBcATS0hFQkDGydJbBsFDwNJbA8NAD8rPysROTk5OSsrKysBLxDOKysrKxE5"
    "OS8vEO0Q7RDtEO0REjk5MDErKysrKysrKysrKysrKysrKysrKysrKysAKysrKysrKysr"
    "Kys3FBYzMjY1NCYnBgYXFAYjIiY1NDY3JiY1NDYzMhYVFAYHFhYnNCYjIgYVFBYXNjYZ"
    "FRATERgYDA1YHBcVHhEQDg8aFxUaDxAREBAQERMQFxUMDC0QFBQPDRgKCBgMExoaEREY"
    "CQcVDREYGBEMFQgKGEoOFRQLCxUIBhIAAAIACwACAHIAsQALACwA+LkAFv/oQCkQAE0O"
    "EBcATSkQFgBNKBgWAE0oGA4SAEwoGAsATRoYCxIATBcYCwBNE7j/8LMWAE0TuP/4QBUO"
    "EABMDRgXAE0NGBESAEwNCAsATQq4/+izFgBNCrj/6LMQAE0KuP/4QAoODwBMBBATAE0C"
    "uP/osxAATQK4//hADw4PAEwGKh0VIR0bAB0PFbj/wLMVAE0PuP/Asw4ATQ+4/8BAJwwA"
    "TQ9ACwBNFS4PKgkDGCdJbAEYHhJzEglJbAEYDANKEgF0EgUYDQA/PysrKysREjkBLxDO"
    "KysrKxDt1O0Q7TIwMSsrKysrKysrKysrKysrKysrACsrNxQWMzI2NTQmIyIGFyImNTQ2"
    "MzIWFRQGIyImNTQ2MzIWFRQWFxYWMzI2NwYGHBENDxYQEhEQGxQYGhgZHCEdDRMFBgYE"
    "AgEBBQMSFgIFFHoWFxwMGB0bSxkcGCAoJio3DAoECAkGBAUBAQEoLwwQAAABAAz/6gD1"
    "AM8AFQAMswkRABEALy8RMzAxNxcHHgIXFQYHLgInDgIHJz4Cbx4KBh0xHhQIGyMWBQQU"
    "MCoCKisPzw0JNUozBQUBDRMzOjMwQTUSAxo9XQAABAAK/+YA9QDPABIALQA5AEIAH0AN"
    "JhoMPzouIhMsLAMuAwAvLxI5LzMzETMyMzIyMDE3FBcHNjU1BgcnNjY3FwcGBxcHFzM3"
    "FwcGBgcmJzUWNjY3Iw4CByc+AjcjBzcXBx4CFxUGByYmJxcHBgYHJzY2QQEQARIVAhYh"
    "CRYJEgoOCCVUCQ0IAgUVBBsbDAQDLAMMHB8BGhgJAQYMOBEIBhccDRIDGR0pFQkQHxQD"
    "Fh8UGQ8GFxJoGRMEGUIgDQUiEwcHFQkNBkUeCQ0IBAQCC04gKB4OBA4iKxUCdwcGHCQU"
    "BAMDCBI+Fg0FJSoQAxVBAAACAA//5gDxANIALgA3AB9ADQ80LxglKyIrGBgcKBwALy8S"
    "OS8zMxEzETMyMjAxNwcWFhUUBiMiJyYnDgIHNCc3FhY2NjcjDgIHJz4CNyMiByczNjcX"
    "BwczNxcHMwYGJjU0NzbCAi4DBgEEBQgaBAoSEiEBGxANCgQ/BBEnJwIhIw8DJQ0LCUYC"
    "ARcHBD0KEZIFBBUNCQ+KHhwOBAgJDRIaNzATBQwPBQgBCzpMJTg1FwMZNjQjAwkeHwwG"
    "KwsOIiwPBAMEBwwAAAQAD//pAPUA0AAIADUATABRACVAEEFNNktLAw81KyEhFxdFA0UA"
    "Ly8SOS8zETMzMxE5LzMzMjAxNzY2NxcGBwYHFzY3FwcGBzM3FwYHJzcjFgYmNTQ3Njcz"
    "FzMmJic3FhYVFAczJiYnNxYWFRQHBzM3FwcGBxY3FQYHJicGByc2NyYnIwc3Fhc2Nyct"
    "YBIQLiUmNnoNChcLEQstCRQSFAMNqQIMDwYKBAQBFwEGCwMPDwguAgYIAw0PCUxyChIM"
    "EhkrQBQCNSswPwE/JxcMDgkcEBgZDrQDDwoUBAMEATQXIA8EFQ8LFAIbAh4VDAMDAwcM"
    "GA0GFQ8CCBAFBwgOEg4CCBIGCAglChEFHBccBAQFCwYfIAgEDh4ZHwICHBUWGwAAAwAQ"
    "/+gA8ADOAA4AEgAwACtAEy0gIA0sKCYmDRAFBQ0LDw8XDRcALy8SOS8zETkvMxE5LzMz"
    "ETkvMzAxNxQXBzUjFQc2NCcXMzcXBxUzNRcOAgcmJzcWMzI2NyMHJzY2NyMHJzM3FyMH"
    "MzcXwAEPZA8BAQ9jCQ56ZBQJDg4QAh0BGw0FDwhcCQwHCQQ5CQjBDRKFDlkLD7stDQYK"
    "CQUULRQICgwEMzOHNBAHBA8MBAgGNwgOBBIOAggNEyQLDwAAAgAL/+YA9QDTAAMALgAx"
    "QBYBIRIeEhYWBCUNAA0REQQMCCgoGgQaAC8vEjkvMzMROS8zMxEzETkvMzMRMzMwMTcV"
    "MzUnFwcGBzM3FyMVMzcXIxUzNxcjFRQXBzY1NSMiByczNCcXMzUjBgYHJzY2Rzo4FgkI"
    "CIIPFVwpDhRLQQ8VZQERAVUNCwktARA6PgoWEwMSIGs0NGgOBAwMDxUyDhQ0DxUmEQ0H"
    "Hw4eAwkuFQkyDxkPAhMzAAMANv/tANIAyAADAAcAFwAnQBEUDgsFFhYLAQQECwwAAA"
    "gLCAAvLxI5LzMROS8zETkvMxEzMjAxNxUzNQcVMzUHNjQnFzM3FwcVFBcHNSMVRXd3d4"
    "YBAQ91BxEIAQ93uExMUldXeTtjPQoKDQaJHRgIGhMAAwAQ/+YAzADKAAMACAAhACNADx8"
    "GDg4aAQQEGhgAABEaEQAvLxI5LzMROS8zETkvMzIwMTcVMzUHFAczNQcWNjU1IwYGBy"
    "c+AycXMzcXBxUUByYnVGFhAWIuGxNiAyUZAhQWCwEBD18JEAkZASK5MjI4FxoxfQQBD"
    "D0oLg4DECArRUAKCw4HshIIDggAAAkADf/nAO8AzwADAAcACwAnAC8AOgA+AEMAXABd"
    "QCxaFA42MCwoIAkiIiAgDkFJSQ4FCAgOPD8/DgEEBA4nERslEQAADlM7O0wOTAAvLxI5"
    "LzMROS8zMzMRMxE5LzMROS8zETkvMxE5LzMROS8zETMRMzIzMhEzMjAxNxUzNQcVMzUH"
    "FTM1JzQnFwcVMzQnFwcVMzcXIxUzNxcjByczNSMHJxcXBwYHJzY2NxYWFRQGIyInJic3"
    "FTM1BxQHMzUHFhY1NSMGBgcnPgI0JxczNxcHFRYHNCc/KSkpKSk3ARcIKQEXCAMLER8B"
    "CxFtDgghCQ4ILBIJGx0CDxkqGAgGAgMDBQ9MMTEBMiQTETICIRkCERYKAQ8vBxAHARc"
    "cnxoaIB4eJB8fShkRDQYXEhgMBhgLEWMLEQIIYwIIcxADIBIDDiAQDQoEBQoIDROKMj"
    "I4HBczfAMBCzwjLA8DDiEsZhoJCgwHrhQIDQgABgAN//QA9wDEABwAQgBGAEoATgBSAF"
    "9ALSMgMTg4Piw7LDAwIFBMK0xBQSAHEgQSFhYgSERPREtLIBENCgogRyFDQzUgNQAvLxI5"
    "LzMzETkvMzMROS8zMxEzETkvMzMRMxE5LzMzETMROS8zMxEzMhEzETMwMTc2NzUjIgcnMz"
    "UjIgcnMzcXIxUzNxcjFTcXBgcHNzY0JxczNxcHFRQXBzUjFTM3FyMVMzcXIyIHJzM1IyIH"
    "JzM1IxU1FTM1MxUzNQcVMzUzFTM1DR4LBQ0LCSYHDQsJPwwRJwoMEScpASYmCFcBAQ1TB"
    "w8IAQ0kHQwROioNEogNCwlUFg0LCTclJQwkVSUMJBcIBEMDCT0DCQwSPQwSPw4EERMIQyA"
    "8IAkJDAYvIBIGDCsMEikNEwMJKQMJKwpoKCgoKC4qKioqAAAEAAz/9QD1AM0AGQArAEYAT"
    "QAnQBFHSikfJiYeGikpPBYTEyM8IwAvLxI5LzMROS8zMzIRMxEzMjAxNwYHHgIGIyInJic"
    "GBgcnNjY3IyIHJzM3FwczNxcjFTM3FyMiByczNSMiBwc3FwYGByc2NjcGBgcnNjY3FwcGBzc"
    "2NxcHBgcnNjcXBgbHDg8hHAMFAgQHCicKIx8BISoSLQ0LCU8KD2NYDBE6LwwRgA0LCUcMDQ"
    "tTQQETJw0MCh0PDxgNCQwhBRUJHBQyDQcUCyAyCjEuARI5sxMTDRAQBwoOGAsaDQQSKiAD"
    "CQoPcAwSRAwSAwlEAwILBAUNCRMDIRcDBggTAzwVDQYuEwMUEg8ELmkSCQsEBhQAAAMACv"
    "/1APUAxQAJABwAOQA1QBgQBhUBHyc0Li4iJiYfGBUVHyEdNzcrHysALy8SOS8zMxE5LzMROS8"
    "zMhEzMxEzETMyMDE3NxYWBgYjIicmFzcXBgcGByc2JzUjIgcnMzcXBzczNxcjFTM3FyMVMzcX"
    "IyIHJzM1NRcHFTM1IyIHIQISDgMHAQQBAwshAxAICQYOBQEJBwsJIgcRCRx2EBNGHg8TQCsP"
    "EZIHCwklFgkjKggLvwIKDg0GCQuXIAIWCwsODwUNWAMJCgwHSQ8VUA4UXA4UAwloHAwHcbI"
    "DAAQAEf/mAO0A0wADAAcADQA1ADFAFg0nJyQkLgUMDC4BBAQuMSsrAAAWLhYALy8SOS8zETM"
    "ROS8zETkvMxE5LzMRMzAxNxUzNQcVMzUHNjc1IxU3NjcXBgcVFAcmJzUWNic1IwYHJzY3IyI"
    "HJzM0JxczNjcXBgczNxcHU2FhYRILB2FvFQYQCSIXAR0ZDwEFSVQBUjpsCAsJMAEPEw0EEw"
    "gWRggQCKgdHSMcHEAKCQseHhoNEQEoUBMHDwYEAgEIOUIRAxs1AwlTHwkYDQwBGAsNBgAA"
    "AAABAAAAGgGBADEAAAAAAAIAEAAvAIgAAAIXAx4AAAAAAAAAAAAwAE0AtgD3AZICSgKiAysD"
    "2gQjBT0F+QYlBp4HAgeTB/MITwiJCNAJhQolCrALHwuJAAAAAQAAAAUmZie1DThfDzz1AAsB"
    "AAAAAAC3mEKAAAAAANOHH7P//v/cAQAA3AAAAAwAAgAAAAAAAAEAAAAAgAAOAIAAEQCAAAwAg"
    "AAdAIAADQCAAA8AgAAJAIAADwCAAAwAgAAQAIAACwCAAAsBAAAMAAoADwAPABAACwA2ABAA"
    "DQANAAwACgARAAEAAADc/9wAJAEA//7/9wEAAAEAAAAAAAAAAAAAAAAAAAAOAAMAgAGQAAUA"
    "CACAAIAAAAAQAIAAgAAAAIAADABBAAACAQYAAwEBAQEBAAAAAQgAAAAAAAAAAAAAAFpZRUMA"
    "QAAsjqsA3P/cACQA3AAkAAAAAQAAAAAAdACvAAAAIAABAAAAAQADAAEAAAAMAAQAmAAAAC"
    "IAIAAEAAIALAAuADlOuk79Up5T11P3XnRl5WcIZx90Bn7Pi8GOq///AAAALAAuADBOuk79Up"
    "5T11P3XnRl5WcIZx90Bn7Pi8GOq////9X/1P/TsVOxEa1xrDmsGqGemi6ZDJj2jBCBSHRXcW4"
    "AAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABARoeGhYSDgoGAf359fHt6eXh"
    "3dnV0c3JxcG9ubWxramloZ2ZlZGNiYWBfXl1cW1pZWFdWVVRTUVBPTk1MS0pJSEdGKB8QCg"
    "ksAbELCkMjQ2UKLSwAsQoLQyNDCy0sAbAGQ7AHQ2UKLSywTysgsEBRWCFLUlhFRBshIVkbIy"
    "GwQLAEJUWwBCVFYWSKY1JYRUQbISFZWS0sALAHQ7AGQwstLEtTI0tRWlggRYpgRBshIVktL"
    "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"
    "ByWwByWwBSWwcSuwAhc4sABSsAIlsAFSWliwBCWwBiVJsAMlsAUlSWAgsEBSWCEbsABSWCCw"
    "AlRYsAQlsAQlsAclsAclSbACFzgbsAQlsAQlsAQlsAYlSbACFzhZWVlZWSEhISEhLSy6AF"
    "0QAAALJWNWYCuwByWwByWwBiWwBiWwDCWwDCWwCSWwCCWwbiuwBBc4sAclsAclsAcmsG0rs"
    "AQlsAQlsAQmsG0rsFArsAYlsAYlsAMlsHErsAUlsAUlsAMlsAIXOCCwBiWwBiWwBSWwcSt"
    "gsAYlsAYlsAQlZbACFziwAiWwAiVgILBAU1ghsEBhI7BAYSMbuP/AUFiwQGAjsEBgI1lZs"
    "AglsAglsAQmsAIXOLAFJbAFJYqwAhc4ILAAUliwBiWwCCVJsAMlsAUlSWAgsEBSWCEbsABS"
    "WLAGJbAGJbAGJbAGJbALJbALJUmwBBc4sAYlsAYlsAYlsAYlsAolsAolsAclsHErsAQXOLA"
    "EJbAEJbAFJbAHJbAFJbBxK7ACFzgbsAQlsAQluf/AAAIXOFlZWSEhISEhISEhLSywBCWwA"
    "yWHsAMlsAMliiCwAFBYIbBlG7BoWStksAQlsAQlBrAEJbAEJUkgIGOwAyUgY1GxAAMlVFt"
    "YISEjIQcbIGOwAiUgY2EgsFMrimOwBSWwBSWHsAQlsAQmSrAAUFhlWbAEJiABRiMARrAF"
    "JiABRiMARrAAFgCwACNIAbAAI0gAILABI0iwAiNIASCwASNIsAIjSCOyAgABCCM4sgIAA"
    "QkjOLECAQewARZZLSwjEA0MimMjimNgZLlAAAQAY1BYsAA4GzxZLSywBiWwCSWwCSWwBya"
    "wdyuwAFBYBRsEWbAEJbAGJrB3K7AFJbAFJrAFJbAFJrB2K7AAUFgFGwRZsHcrLSywAyWwA"
    "iWwAiWwCCawdiqKILAAVFgFGwRZsAQmsAYmsHcrsAUmsAUmsAUmsAUmsHYrsABUWAQbBFmw"
    "dyssI4pKILAAVFgFGwRZsAMlsAcmsHcrsAIlsAYmsHcrsAMlsAUmsHcrIrAAVFgFGwRZsA"
    "MlsAgmsHcrsAUmsAUmsHYrIrAAVFgFGwRZsAMlsAgmsHcrsAUmsAUmsHcrLSywAyWwAiWw"
    "AiWwCCawdiuwA1gigFNYsABQWCGwZhuwaFkrsAMlsAUmsHcrI4pKIrAAVFgFGwRZsAMlsA"
    "clsHcrsAUmsAUmsHYrILAAVFgFGwRZsAMlsAcmsHcrsAUmsAUmsHcrLSywAyWwAiWwAiW"
    "wCCawdiuwCCawdisjiiCwA1RYsABQWCGwZhuwaFkrsAMlsAUmsHcrI4pKIrAAVFgFGwRZ"
    "sAMlsAclsHcrsAUmsAUmsHYrILAAVFgFGwRZsAMlsAcmsHcrsAUmsAUmsHcrLSywBCawB"
    "CawBCawdisjiiCwA1RYsABQWCGwZhuwaFkrsAQmsAUmsHcrI4pKIrAAVFgFGwRZsAQmsA"
    "clsHcrsAUmsAUmsHYrILAAVFgFGwRZsAQlsAcmsHcrsAUmsAUmsHcr"
)

def extract_font_data(doc):
    """从PDF提取SimSun字体数据"""
    for xref in range(1, doc.xref_length()):
        try:
            obj = doc.xref_object(xref)
            if not obj:
                continue
            if "/FontFile2" in obj:
                match = re.search(r'/FontFile2\s+(\d+)', obj)
                if match:
                    ff2_xref = int(match.group(1))
                    data = doc.xref_stream(ff2_xref)
                    if data and len(data) > 1000:
                        return data
        except:
            continue
    for xref in range(2, min(20, doc.xref_length())):
        try:
            data = doc.xref_stream(xref)
            if data and len(data) > 10000:
                return data
        except:
            continue
    return None


def make_comma_font():
    """从内嵌base64生成逗号字体文件"""
    subset_path = '/tmp/simsun_subset_comma_v14.ttf'
    if os.path.exists(subset_path) and os.path.getsize(subset_path) > 1000:
        return subset_path
    
    try:
        font_data = base64.b64decode(COMMA_FONT_B64)
        with open(subset_path, 'wb') as f:
            f.write(font_data)
        st.info(f"✅ 内嵌字体释放成功: {len(font_data)} bytes")
        return subset_path
    except Exception as e:
        st.error(f"❌ 内嵌字体释放失败: {e}")
        return None


def fmt_decimal(value, field_key):
    """格式化数字"""
    if not value or not str(value).strip():
        return value
    if field_key.startswith("eq"):
        return str(value).strip()
    if field_key in ["agent_name", "agent_id", "receiver", "receive_date"]:
        return str(value).strip()
    text = str(value).strip().replace(",", "")
    if "." in text or "%" in text:
        return text
    try:
        return "{:,.2f}".format(int(text))
    except:
        return text


FIELD_CFG = {
    # --- 第1行: 从业人数 ---
    "eq1s": (0, 100.6, 152.3, 164.8, 183.6, 0.5),
    "eq1e": (0, 152.6, 204.0, 164.8, 183.6, 0.5),
    "eq2s": (0, 204.3, 255.7, 164.8, 183.6, 0.5),
    "eq2e": (0, 256.0, 307.4, 164.8, 183.6, 0.5),
    "eq3e": (0, 514.5, 565.9, 164.8, 183.6, 0.5),
    # --- 第2行: 资产总额 ---
    "aq1s": (0, 100.6, 152.3, 183.6, 202.3, 0.5),
    "aq1e": (0, 152.6, 204.0, 183.6, 202.3, 0.5),
    "aq2s": (0, 204.3, 255.7, 183.6, 202.3, 0.5),
    "aq2e": (0, 256.0, 307.4, 183.6, 202.3, 0.5),
    "aq3e": (0, 514.5, 565.9, 183.6, 202.3, 0.5),
    # --- 预缴税款计算 ---
    "L1": (0, 514.5, 565.9, 297.7, 305.7, 0.2),
    "L2": (0, 514.5, 565.9, 316.5, 324.5, 0.2),
    "L3": (0, 514.5, 565.9, 335.2, 343.2, 0.2),
    "L4": (0, 514.5, 565.9, 354.0, 362.0, 0.2),
    "L5": (0, 514.5, 565.9, 372.7, 380.7, 0.2),
    "L6": (0, 514.5, 565.9, 391.5, 399.5, 0.2),
    "L7": (0, 514.5, 565.9, 410.2, 418.2, 0.2),
    "L8": (0, 514.5, 565.9, 429.0, 437.0, 0.2),
    "L9": (0, 514.5, 565.9, 447.7, 455.7, 0.2),
    "L10": (0, 514.5, 565.9, 466.5, 474.5, 0.2),
    "L11": (0, 514.5, 565.9, 485.2, 493.2, 0.2),
    "L12": (0, 514.5, 565.9, 504.0, 512.0, 0.2),
    "L13": (0, 514.5, 565.9, 522.7, 530.7, 0.2),
    "L13_1": (0, 514.5, 565.9, 541.5, 549.5, 0.2),
    "L14": (0, 514.5, 565.9, 560.2, 568.2, 0.2),
    "L15": (0, 514.5, 565.9, 579.0, 587.0, 0.2),
    "L16": (0, 514.5, 565.9, 597.7, 605.7, 0.2),
    "L17": (0, 514.5, 565.9, 631.5, 639.5, 0.2),
    "L18": (0, 514.5, 565.9, 650.2, 658.2, 0.2),
    "L19": (0, 514.5, 565.9, 669.0, 677.0, 0.2),
    "L20": (0, 514.5, 565.9, 687.7, 695.7, 0.2),
    "L21": (0, 514.5, 565.9, 706.5, 714.5, 0.2),
    "L22": (0, 514.5, 565.9, 725.2, 733.2, 0.2),
    "FZ1": (0, 514.5, 565.9, 762.7, 770.7, 0.2),
    "FZ2": (0, 514.5, 565.9, 781.5, 789.5, 0.2),
    "L23": (0, 514.5, 565.9, 800.0, 808.0, 0.2),
    # --- 第2页 ---
    "L23_2": (1, 514.5, 565.9, 10.0, 27.6, 0.2),
    "FZ3": (1, 514.5, 565.9, 43.0, 51.0, 0.2),
    "L24": (1, 514.5, 565.9, 61.7, 69.7, 0.2),
    # --- 第2页签章 ---
    "agent_name": (1, 80.8, 184.8, 103.4, 111.4, 2.0),
    "agent_id": (1, 112.8, 264.8, 113.0, 121.0, 2.0),
    "receiver": (1, 355.7, 463.7, 103.4, 111.4, 2.0),
    "receive_date": (1, 355.7, 483.7, 122.6, 130.6, 2.0),
    # --- 第3页 ---
    "A201_R1C1": (2, 224.0, 288.0, 104.0, 112.0, 0.2),
    "A201_R1C2": (2, 288.0, 352.0, 104.0, 112.0, 0.2),
    "A201_R1C3": (2, 352.0, 416.0, 104.0, 112.0, 0.2),
    "A201_R1C4": (2, 416.0, 480.0, 104.0, 112.0, 0.2),
    "A201_R1C5": (2, 480.0, 566.0, 104.0, 112.0, 0.2),
    "A201_R2C1": (2, 224.0, 288.0, 120.7, 128.7, 0.2),
    "A201_R2C2": (2, 288.0, 352.0, 120.7, 128.7, 0.2),
    "A201_R2C3": (2, 352.0, 416.0, 120.7, 128.7, 0.2),
    "A201_R2C4": (2, 416.0, 480.0, 120.7, 128.7, 0.2),
    "A201_R2C5": (2, 480.0, 566.0, 120.7, 128.7, 0.2),
    "A201_R3C1": (2, 224.0, 288.0, 135.7, 143.7, 0.2),
    "A201_R3C2": (2, 288.0, 352.0, 135.7, 143.7, 0.2),
    "A201_R3C3": (2, 352.0, 416.0, 135.7, 143.7, 0.2),
    "A201_R3C4": (2, 416.0, 480.0, 135.7, 143.7, 0.2),
    "A201_R3C5": (2, 480.0, 566.0, 135.7, 143.7, 0.2),
}


def fill_and_render(pdf_bytes, values, dpi=300, fmt="png"):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    font_data = extract_font_data(doc)
    font_simsun = None
    if font_data:
        try:
            font_simsun = fitz.Font(fontbuffer=font_data)
            st.info(f"✅ 提取字体成功: {len(font_data)} bytes")
        except Exception as e:
            st.warning(f"⚠️ 提取字体失败: {e}")
    
    comma_font_path = make_comma_font()
    font_comma = None
    if comma_font_path:
        try:
            font_comma = fitz.Font(fontfile=comma_font_path)
            st.info("✅ 逗号字体载入成功")
        except Exception as e:
            st.warning(f"⚠️ 逗号字体载入失败: {e}")
    
    use_fallback = (font_simsun is None)
    if use_fallback:
        st.warning("⚠️ 使用备份字体 china-ss")
    
    try:
        for key, raw_value in values.items():
            if not raw_value or key not in FIELD_CFG:
                continue
            
            new_value = fmt_decimal(raw_value, key)
            if not new_value:
                continue
            
            page_num, x0, x1, y0, y1, right_margin = FIELD_CFG[key]
            page = doc[page_num]
            text = str(new_value)
            
            # 白色覆盖旧文字
            INSET = 1.0
            for b in page.get_text("dict")["blocks"]:
                if "lines" not in b:
                    continue
                for line in b["lines"]:
                    for span in line["spans"]:
                        sb = span["bbox"]
                        if (sb[0] >= x0 - 1 and sb[2] <= x1 + 1 and
                            sb[1] >= y0 - 1 and sb[3] <= y1 + 1):
                            cl = max(sb[0] + INSET, x0 + INSET)
                            cr = min(sb[2] - INSET, x1 - INSET)
                            ct = max(sb[1] + INSET, y0 + INSET)
                            cb = min(sb[3] - INSET, y1 - INSET)
                            if cr > cl and cb > ct:
                                rect = fitz.Rect(cl, ct, cr, cb)
                                shape = page.new_shape()
                                shape.draw_rect(rect)
                                shape.finish(color=(1, 1, 1), fill=(1, 1, 1))
                                shape.commit()
            
            # 写入新文字
            origin_y = y0 + (y1 - y0) * 0.75
            
            if key in ["agent_name", "agent_id", "receiver", "receive_date"]:
                write_x = x0 + 2.0
                if use_fallback:
                    page.insert_text((write_x, origin_y), text, fontname="china-ss",
                                     fontsize=FONTSIZE, color=(0, 0, 0))
                else:
                    tw = fitz.TextWriter(page.rect)
                    tw.append((write_x, origin_y), text, fontsize=FONTSIZE, font=font_simsun)
                    tw.write_text(page, color=(0, 0, 0))
            else:
                # 数值字段: 逐字符右对齐
                total_width = 0
                for char in text:
                    if char == ',':
                        total_width += COMMA_W
                    elif char == '.':
                        total_width += DOT_W
                    else:
                        total_width += CHAR_W
                
                current_x = (x1 - right_margin) - total_width
                
                if use_fallback:
                    for char in text:
                        page.insert_text((current_x, origin_y), char, fontname="china-ss",
                                         fontsize=FONTSIZE, color=(0, 0, 0))
                        if char == ',':
                            current_x += COMMA_W
                        elif char == '.':
                            current_x += DOT_W
                        else:
                            current_x += CHAR_W
                else:
                    tw = fitz.TextWriter(page.rect)
                    for char in text:
                        if char in [',', '.'] and font_comma is not None:
                            tw.append((current_x, origin_y), char, fontsize=FONTSIZE, font=font_comma)
                        else:
                            tw.append((current_x, origin_y), char, fontsize=FONTSIZE, font=font_simsun)
                        
                        if char == ',':
                            current_x += COMMA_W
                        elif char == '.':
                            current_x += DOT_W
                        else:
                            current_x += CHAR_W
                    tw.write_text(page, color=(0, 0, 0))
        
        images = []
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            if fmt == "png":
                img_bytes = pix.tobytes("png")
            else:
                img_bytes = pix.tobytes("jpeg")
            images.append(img_bytes)
        
        return images
    finally:
        doc.close()


def main():
    st.title("PDF智能填表系统 v14.2")
    st.markdown("TextWriter + fitz.Font | **零外部依赖版** | PNG输出")

    st.header("1️⃣ 上传PDF模板")
    uploaded_file = st.file_uploader("选择PDF文件", type=["pdf"])

    if uploaded_file is None:
        st.info("👆 请先上传PDF模板文件")
        return

    pdf_bytes = uploaded_file.getvalue()

    st.header("2️⃣ 填写字段数据")
    st.caption("留空表示不修改。数字自动添加两位小数")

    values = {}

    with st.expander("👥 第1行 - 从业人数", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(总)"],
                                   ["eq1s","eq1e","eq2s","eq2e","eq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)

    with st.expander("💰 第2行 - 资产总额", expanded=True):
        cols = st.columns(5)
        for col, label, key in zip(cols, ["Q1季初","Q1季末","Q2季初","Q2季末","Q3季末(总)"],
                                   ["aq1s","aq1e","aq2s","aq2e","aq3e"]):
            with col: values[key] = st.text_input(label, value="", key=key)

    with st.expander("📊 第3-16行（预缴税款计算）", expanded=True):
        col1, col2 = st.columns(2)
        left = [("L1","3 营业收入"),("L2","4 营业成本"),("L3","5 利润总额"),("L4","6 特定业务"),
                ("L5","7 不征税收入"),("L6","8 减：资产加速折旧"),("L7","9 减：免税收入")]
        right = [("L8","10 减：所得减免"),("L9","11 减：所得减免其他"),("L10","12 实际利润额"),
                 ("L11","13 税率(25%)"),("L12","14 应纳所得税额"),("L13","15 减免所得税额"),("L13_1","15.1 减免明细")]
        with col1:
            for k,l in left: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right: values[k] = st.text_input(l, value="", key=k)

    with st.expander("📋 第16-25行"):
        col1, col2 = st.columns(2)
        left2 = [("L14","16 本期预缴"),("L15","17 减免其他"),("L16","18 本期应补(退)"),
                 ("L17","19 总机构本期"),("L18","20 总机构分摊"),("L19","21 财政集中")]
        right2 = [("L20","22 分支机构"),("L21","23 分摊比例"),("L22","24 税率"),
                  ("FZ1","附1 中央级收入"),("FZ2","附2 地方级收入"),("L23","25 减免地方")]
        with col1:
            for k,l in left2: values[k] = st.text_input(l, value="", key=k)
        with col2:
            for k,l in right2: values[k] = st.text_input(l, value="", key=k)

    with st.expander("📄 第2页 - 补充字段"):
        col1, col2, col3 = st.columns(3)
        with col1: values["L23_2"] = st.text_input("23.2 本年累计应减免", value="", key="L23_2")
        with col2: values["FZ3"] = st.text_input("FZ3 地方级收入实际应纳税额", value="", key="FZ3")
        with col3: values["L24"] = st.text_input("24 实际应补(退)所得税额", value="", key="L24")

    with st.expander("✏️ 第2页 - 签章信息"):
        c1, c2 = st.columns(2)
        with c1:
            values["agent_name"] = st.text_input("经办人", value="", key="agent_name")
            values["agent_id"] = st.text_input("经办人身份证号", value="", key="agent_id")
        with c2:
            values["receiver"] = st.text_input("受理人", value="", key="receiver")
            values["receive_date"] = st.text_input("受理日期", value="", key="receive_date")

    with st.expander("📎 第3页 - A201020 资产加速折旧附表"):
        st.caption("行1: 加速折旧")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R1C1"] = st.text_input("R1 账载折旧", value="", key="A201_R1C1")
        with c2: values["A201_R1C2"] = st.text_input("R1 税收一般规定", value="", key="A201_R1C2")
        with c3: values["A201_R1C3"] = st.text_input("R1 加速政策计算", value="", key="A201_R1C3")
        with c4: values["A201_R1C4"] = st.text_input("R1 纳税调减金额", value="", key="A201_R1C4")
        with c5: values["A201_R1C5"] = st.text_input("R1 加速优惠金额", value="", key="A201_R1C5")
        st.caption("行2: 一次性扣除")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R2C1"] = st.text_input("R2 账载折旧", value="", key="A201_R2C1")
        with c2: values["A201_R2C2"] = st.text_input("R2 税收一般规定", value="", key="A201_R2C2")
        with c3: values["A201_R2C3"] = st.text_input("R2 加速政策计算", value="", key="A201_R2C3")
        with c4: values["A201_R2C4"] = st.text_input("R2 纳税调减金额", value="", key="A201_R2C4")
        with c5: values["A201_R2C5"] = st.text_input("R2 加速优惠金额", value="", key="A201_R2C5")
        st.caption("行3: 合计")
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: values["A201_R3C1"] = st.text_input("R3 账载折旧", value="", key="A201_R3C1")
        with c2: values["A201_R3C2"] = st.text_input("R3 税收一般规定", value="", key="A201_R3C2")
        with c3: values["A201_R3C3"] = st.text_input("R3 加速政策计算", value="", key="A201_R3C3")
        with c4: values["A201_R3C4"] = st.text_input("R3 纳税调减金额", value="", key="A201_R3C4")
        with c5: values["A201_R3C5"] = st.text_input("R3 加速优惠金额", value="", key="A201_R3C5")

    st.header("3️⃣ 输出设置")
    c1, c2 = st.columns(2)
    with c1:
        output_fmt = st.radio("图片格式", ["PNG（无损推荐）", "JPG"], horizontal=True)
    with c2:
        dpi = st.select_slider("DPI", options=[150, 200, 300, 400], value=300)
    fmt = "png" if "PNG" in output_fmt else "jpeg"
    ext = "png" if fmt == "png" else "jpg"

    filled = {k: v.strip() for k, v in values.items() if v.strip()}
    if filled:
        st.success(f"✅ 已填写 {len(filled)} 个字段")
    else:
        st.info("💡 尚未填写任何字段")

    if st.button("🚀 生成图片", type="primary", disabled=len(filled)==0):
        with st.spinner("正在渲染..."):
            try:
                images = fill_and_render(pdf_bytes, filled, dpi=dpi, fmt=fmt)
                st.success(f"✅ 成功生成 {len(images)} 页图片（{dpi} DPI）")
                if len(images) == 1:
                    st.download_button(f"📥 下载图片（{ext.upper()}）", images[0],
                        f"filled_tax_form_page1.{ext}", mime=f"image/{ext}", use_container_width=True)
                else:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for i, img_data in enumerate(images):
                            zf.writestr(f"filled_tax_form_page{i+1}.{ext}", img_data)
                    zip_buf.seek(0)
                    st.download_button(f"📥 下载全部（ZIP）", zip_buf.getvalue(),
                        "filled_tax_form.zip", mime="application/zip", use_container_width=True)
                st.subheader("预览（第1页）")
                st.image(images[0], use_container_width=True)
            except Exception as e:
                st.error(f"❌ 生成失败: {e}")
                st.exception(e)

    st.markdown("---")
    st.markdown("<center>PDF智能填表系统 v14.2 | TextWriter + fitz.Font | 零外部依赖版</center>",
                unsafe_allow_html=True)


if __name__ == "__main__":
    main()

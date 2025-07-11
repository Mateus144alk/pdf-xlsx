from pytube import YouTube

def baixar_video(url, caminho_saida='.'):
    try:
        yt = YouTube(url)
        stream = yt.streams.get_highest_resolution()
        print(f"Baixando: {yt.title}")
        stream.download(output_path=caminho_saida)
        print("Download concluído com sucesso!")
    except Exception as e:
        print("Ocorreu um erro:", e)

# Exemplo de uso
if __name__ == "__main__":
    url_video = input("Cole a URL do vídeo do YouTube: ")
    baixar_video(url_video)

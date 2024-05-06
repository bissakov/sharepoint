from src.sharepoint import SharePoint
import pathlib
import dotenv
import rich
import time
import os


def main():
    dotenv.load_dotenv()

    project_dir = pathlib.Path(os.path.abspath(__file__)).parents[1].as_posix()

    folder_url = "/Shared Documents"

    client_id = os.getenv("CLIENT_ID")
    assert client_id is not None, "CLIENT_ID is not set"

    client_secret = os.getenv("CLIENT_SECRET")
    assert client_secret is not None, "CLIENT_SECRET is not set"

    base_url = os.getenv("BASE_URL")
    assert base_url is not None, "BASE_URL is not set"

    sharepoint = SharePoint(base_url, client_id, client_secret)

    # rich.print(sharepoint.list_files(folder_url))
    # rich.print(sharepoint.list_files(folder_url, include_properties=True))
    #
    # rich.print("##################################################")
    #
    # rich.print(sharepoint.list_subfolders(folder_url))
    # rich.print(sharepoint.list_subfolders(folder_url, include_properties=True))

    # rich.print(sharepoint.list_folder_contents(folder_url, recursive=True))
    # rich.print(sharepoint.list_folder_contents(folder_url, recursive=True))
    # time.sleep(10)
    # rich.print(sharepoint.list_folder_contents(folder_url, recursive=True))
    # time.sleep(5)
    # rich.print(sharepoint.list_folder_contents(folder_url, recursive=True))
    #
    # temp_dir = pathlib.Path(project_dir, "temp").as_posix()
    # temp_zip = pathlib.Path(temp_dir, "temp.zip").as_posix()
    # download_path = pathlib.Path(temp_dir, "Rekvizity (1).docx").as_posix()
    # file_to_download = pathlib.Path(folder_url, "Rekvizity (1).docx").as_posix()
    #
    # sharepoint.read_file(file_to_download)
    # sharepoint.download_file(file_to_download, download_path)
    #
    # sharepoint.download_folder(
    #     folder_url, temp_dir, download_option="all", recursive=True
    # )
    #
    # sharepoint.download_folder(folder_url, temp_zip, recursive=True)
    # sharepoint.create_folder(folder_url, "Test_03-05-2024")
    # sharepoint.upload_file(
    #     remote_folder_url="/Shared Documents/Test_03-05-2024",
    #     local_file_path=r"D:\Work\python_rpa\sharepoint\temp\wget-1.11.4-1-bin.zip",
    # )
    # sharepoint.delete_file(
    #     file_url="/Shared Documents/Test_03-05-2024/wget-1.11.4-1-bin.zip"
    # )
    # sharepoint.delete_folder(folder_url="/Shared Documents/Test_03-05-2024")

    # sharepoint.upload_folder(
    #     remote_folder_url="/Shared Documents/Test_03-05-2024",
    #     local_folder=r"D:\Work\python_rpa\sharepoint\temp\Shared Documents",
    # )

    # sharepoint.upload_folder_as_zip(
    #     remote_folder_url="/Shared Documents/Test_03-05-2024",
    #     local_folder_path=r"D:\Work\python_rpa\sharepoint\temp\Shared Documents",
    # )

    test_folder_url = "/Shared Documents/Test_03-05-2024"

    # sharepoint.upload_file(
    #     remote_folder_url=test_folder_url,
    #     local_file_path=r"D:\Downloads\Telegram Desktop\PythonRPAStudio.setup.0.1.79 (2).exe",
    #     chunk_size_bytes=262_144_000,
    # )

    sharepoint.upload_file(
        remote_folder_url=test_folder_url,
        local_file_path=r"D:\Downloads\Telegram Desktop\7260047342874411010.json",
    )

    # sharepoint.create_list(
    #     list_name="TestList", description="Test list", template_name="GenericList"
    # )
    # sharepoint.create_list(
    #     list_name="TestList", description="Test list", template_name="dasldkja"
    # )
    # sharepoint.update_list_name(
    #     original_list_name="TestListawsdkj", new_list_name="TestList"
    # )


if __name__ == "__main__":
    main()

from odt_to_pdf import convert_to_pdf

import config
import logging

from aiogram import Bot, Dispatcher, executor, types

logging.basicConfig(level=logging.INFO)


bot = Bot(token=config.TOKEN)
dp = Dispatcher(bot)


@dp.message_handler(content_types=types.ContentType.DOCUMENT)
async def handle_docs(message):
    try:
        file_id = message.document.file_id
        downloaded_file = await bot.get_file(file_id)
        file_name = message.document.file_name
        path_to_converted_file = 'files/converted/{file_name}.pdf'.format(file_name=file_name)
        src = '{root_path}/files/converted/{file_name}'.format(root_path=config.ROOT_PATH, file_name=file_name)
        await downloaded_file.download(src)
        response = convert_to_pdf(path_to_odt_document=src)

        with open(path_to_converted_file, 'wb') as f:
            f.write(response)

        # message.answer_document(open(path_to_converted_file, "RB"))
        doc = open(path_to_converted_file, 'rb')
        await message.reply_document(doc)

    except Exception as e:
        print('failed', e.__traceback__)


if __name__ == "__main__":
    executor.start_polling(dp, skip_updates=True)

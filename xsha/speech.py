from aip import AipSpeech

APP_ID = '11468048'
API_KEY = 'BF0OERtklEYDGG9k6mGh5cL4'
SECRET_KEY = 'alw8VKwzadfbWfP5T1iM6xVjxFimfLXd'


def tts(strings, file=None):
    client = AipSpeech(APP_ID, API_KEY, SECRET_KEY)
    result = client.synthesis(strings)
    if file and not isinstance(result, dict):
        file.write(result)
    else:
        return result

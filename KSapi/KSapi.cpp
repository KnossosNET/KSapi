#include <sapi.h>
#pragma warning(disable:4996) 
#include <sphelper.h>
#pragma warning(default: 4996)

ISpObjectToken* speech_set_voice(unsigned int voice_index);

int main(int argc, char* argv[])
{
    unsigned int voice_index = 0;
    int volume = 100;
    int text_index = -1;

    if (argc >= 2) 
    {
        for (int i = 1; i < argc; i+=2) 
        {
            if (strcmp(argv[i], "-text") == 0)
            {
                text_index = i + 1;
            }
            if (strcmp(argv[i], "-vol") == 0)
            {
                volume = atoi(argv[i + 1]);                
            }
            if (strcmp(argv[i], "-voice") == 0)
            {
                voice_index = atoi(argv[i + 1]);
            }
        }
    }
    else
    {
        return 0;
    }

    if (text_index == -1)
    {
        return -1;
    }

    ISpVoice* pVoice = NULL;

    if (FAILED(::CoInitialize(NULL)))
    {
        return -1; 
    }

    HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void**)&pVoice);

    if (SUCCEEDED(hr))
    {
        ISpObjectToken* voice_token = speech_set_voice(voice_index);
        if (voice_token != NULL)
        {
            pVoice->SetVoice(voice_token);
        }
        if (volume >= 0 && volume <= 100)
        {
            pVoice->SetVolume(volume);
        }
        else
        {
            pVoice->SetVolume(100);
        }
        pVoice->Speak(CA2CT(argv[text_index]), 0, NULL);
        pVoice->Release();
        pVoice = NULL;
    }
    else
    {
        return -1;
    }

    ::CoUninitialize();
    return 0;
}

ISpObjectToken* speech_set_voice(unsigned int voice_index)
{
    IEnumSpObjectTokens* cpEnum;
    HRESULT hr = SpEnumTokens(SPCAT_VOICES, NULL, NULL, &cpEnum);
    ULONG ulCount = 0;
    hr = cpEnum->GetCount(&ulCount);

    ISpObjectToken* token;
    if (SUCCEEDED(hr) && voice_index < ulCount)
    {
        cpEnum->Item(voice_index,&token);
        cpEnum->Release();
        return token;
    }
    return NULL;
}
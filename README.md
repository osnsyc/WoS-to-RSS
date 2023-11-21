## Git clone

```shell
git clone https://github.com/osnsyc/WoS-to-RSS.git
cd WoS-to-RSS
```

```python
pip install beautifulsoup4 DrissionPage PyVirtualDisplay Requests translators xlrd
```

## Config 

```shell
vim config.ini
```

```ini
# config.ini
[ID]
; Web of science account
EMAIL = NAME@MAIL.COM
EMAIL_PASSWORD = EMAIL_PASSWORD

; False - check by carsi
IN_SCHOOL = True
; carsi setting
UNIVERSITY = xx大学
STUDENT_ID = 122333
STUDENT_PASSWORD = 122333444555

[Translator]
; translator engine: https://github.com/UlionTse/translators
; disabled - disable translator
TRANSLATOR = baidu

```

## Run

```python
python wos_to_rss.py
```

```python
python wos_server.py
```

## RSS订阅

`http://YOUR_HOST:9277/wos.xml`
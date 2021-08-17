## MOEX-rate-script
The script goes to the site https://www.moex.com/ using selenium.  
Gets the dollar and euro exchange rates for the current month in xml.  
Parses xml, writes data to excel file.  
In excel, it aligns columns to the width, and sets the financial format for the currency.  
After that, it sends the file to the mail (from .env file),  
writes the number of lines of the file in the body of the letter, uses the correct declension.  

### Instalation
1. install requirements
<pre><code>pip install -r requirements.txt</code></pre>

2. сreate a .env file with mail credentials  
<pre><code>sudo vim .env</code></pre>
the following content  
<pre><code>
SMTP_SERVER=your_smtp_server
SMTP_PASSWORD=your_smtp_server_password
RECEIVER=receiver_mail
</code></pre>
If you use Google Mail, you need to enable access for third-party applications in the security settings.  
3. to run
<pre><code>python3 rate.py</code></pre>
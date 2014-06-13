# Command Windows

An ASP homepage styled after the Windows command prompt. Commands such as `help`, `clear`, and `open [title]` typed into the input line on the black screen are handled by the server and their results are appended right below, and `open` shows a program introduction post inside the terminal screen. Clicking a picture in the left menu fills the input line with the command that opens that post. It was made as a school assignment.

<p>
  <img src="docs/screenshots/screenshot-1.jpg" width="480" alt="Opened in Chrome">
</p>


## How to use

The site runs on IIS with ASP enabled. The Releases download contains the site folder `command-windows` along with a readme and the license.

1. In Control Panel, open "Turn Windows features on or off" and enable ASP under "Internet Information Services" > "World Wide Web Services" > "Application Development Features". Screenshots of the setup are in `source/[]HowToSetupASP_1.png` and `source/[]HowToSetupASP_2.png`.
2. Put the site folder under `C:\inetpub\wwwroot\`.
3. Open `http://localhost/<folder name>/` in a browser. From another computer, use the server's IP address instead of `localhost`.

The screen uses the Courier New font and was made for Google Chrome. On first visit, the help text is already shown as if `help` had been typed. Commands are not case sensitive.

| Command | Action |
|---|---|
| `help` | Shows the usage text |
| `clear` | Clears the screen |
| `open <title>` | Opens a post. The titles are `3DP 칩`, `반디집`, `꿀뷰`, `알송`, and `팟플레이어` |
| `<command> & clear` | Clears the screen, then runs the command |

Any other input gets a message saying it is not a command. Clicking HOME in the left menu puts `help & clear` into the input line, and clicking a program picture puts `open <title> & clear`, which runs when Enter is pressed.


## How it works

**The screen history is kept in the session.** The input line is a form that POSTs to `bottom_main.asp` itself, and after handling the command, the ASP code builds the prompt, the command, and the result box as HTML and appends them to `Session("log")`. The page redraws this whole history every time and places a new input line at the bottom, so earlier commands and results stay above even though the page is reloaded, just like a terminal. `clear` empties this session value, and a trailing ` & clear` is stripped from the command, the history is emptied, and the rest of the command is handled.

**Posts are read from files and inserted.** Each post is an HTML fragment file named `bottom_<name>.html`, and when an `open` command comes in, `ADODB.Stream` reads that file as UTF-8 and puts it into the result box as it is.

```asp
Set stream = Server.CreateObject("ADODB.Stream")
With Stream
    .Type = 2
    .CharSet = "UTF-8"
    .Open
    .LoadFromFile(Server.MapPath(".\" & filename))
    post = .ReadText()
    .Close
End With
```

**The screen is split into frames.** The frameset in `index.html` divides the page into the top title bar (`top_main.html`), the left menu (`left_main.html`), the terminal on the right (`bottom_main.asp`), and the bottom border (`bottom_draw.html`). The pictures in the left menu are forms with `target="bottom"`, so they send their commands to the terminal frame. The scrollbars are replaced with nanoScrollerJS, and Rainbow Links makes the links in posts change color on mouseover.


## Files

| Path | Content |
|---|---|
| `source/index.html` | Frame layout |
| `source/top_main.html` | Top title bar |
| `source/left_main.html` | Left menu |
| `source/bottom_main.asp` | Command handling and the terminal screen |
| `source/bottom_3dpchip.html` and four others | Posts opened with `open` |
| `source/bottom_draw.html` | Bottom border |
| `source/image/` | Menu pictures and icons |
| `source/scrollbar_main.js` | Scrollbar setup |
| `source/scrollbar_jquery.min.js`, `source/scrollbar.js`, `source/scrollbar.css`, `source/rainbow.js` | Scripts made by other people |
| `source/[]ReadMe.txt`, `source/[]HowToSetupASP_*.png` | Readme and ASP setup pictures from the original distribution |
| `docs/screenshots/` | Screenshots |
| Releases | The whole site |


## Credits

- `scrollbar_jquery.min.js` is jQuery 1.10.2 by the jQuery Foundation, under the MIT license.
- `scrollbar.js` and `scrollbar.css` are nanoScrollerJS 0.8.0 by James Florentino, under the MIT license.
- `rainbow.js` is Rainbow Links 1.03 by TAKANASHI Mizuki.
- The program banners and icons in `image/` belong to the companies that make those programs.


## License

zlib. See [LICENSE](LICENSE). Bundled libraries made by other people keep their own licenses.

#MakeFile C,C++用 SCodeにてソースファイルの拡張子選択
SCode	 := cpp

CC        = g++
CXXFLAGS  = -g -MMD -Wall -pipe -finput-charset=UTF-8 -fexec-charset=CP932
LDFLAGS   = -lwinmm
LIBS      =
INCLUDE   = -I/usr/include -I/mingw64/include -I./include
TARGETS   = Project.exe
TARGETDIR = .
SRCROOT   = src
OBJROOT   = obj

RM=rm
MENU=windres

# ソースディレクトリの中を(shellの)findコマンドで走破してサブディレクトリまでリスト化する
SRCDIRS  := $(shell find $(SRCROOT) -type d)
# ソースディレクトリを元にforeach命令で全cppファイルをリスト化する
SOURCES   = $(foreach dir, $(SRCDIRS), $(wildcard $(dir)/*.${SCode}))#$()だとエラーでた
# 上記のcppファイルのリストを元にオブジェクトファイル名を決定
OBJECTS   = $(subst $(SRCROOT),$(OBJROOT), $(SOURCES:.${SCode}=.o))
# ソースディレクトリと同じ構造で中間バイナリファイルの出力ディレクトリをリスト化
OBJDIRS   = $(addprefix $(OBJROOT)/, $(SRCDIRS))
# 依存ファイル.dを.oファイルから作る
DEPENDS  := $(OBJECTS:.o=.d)

# 依存ファイルを元に実行ファイルを作る
$(TARGETS): $(OBJECTS) $(LIBS)
	$(CC) -o $(TARGETDIR)/$@ $^ $(LDFLAGS)
	@echo FINISHED!

# 中間バイナリのディレクトリを掘りながら.cppを中間ファイル.oに
$(OBJROOT)/%.o: $(SRCROOT)/%.${SCode}
	@if [ ! -e `dirname $@` ]; then mkdir -p `dirname $@`; fi
	$(CC) $(CXXFLAGS) $(INCLUDE) -o $@ -c $<

# 依存関係ファイル
-include $(DEPENDS)

clean:
	$(RM) $(OBJECTS) $(TARGETS) $(DEPENDS)

.PHONY: all clean
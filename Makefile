# A simple makefile for creating the Bubbling Fluidized Bed Reactor Model distribution
VERSION    := 2015.10.00
PRODUCT    := Bubbling Fluidized Bed Reactor Model
PROD_SNAME := BFBReactor
LICENSE    := CCSI_TE_LICENSE_$(PROD_SNAME).txt
PKG_DIR    := CCSI_$(PROD_SNAME)_$(VERSION)
PACKAGE    := $(PKG_DIR).zip

# Where Jenkins should checkout ^/projects/common/trunk/
COMMON     := .ccsi_common
LEGAL_DOCS := LEGAL \
           CCSI_TE_LICENSE.txt

PAYLOAD := BFB_ACM \
     BFB_gPROMS \
     LEGAL \
     $(LICENSE)

# Get just the top part (not dirname) of each entry so cp -r does the right thing
PAYLOAD_TOPS := $(sort $(foreach v,$(PAYLOAD),$(shell echo $v | cut -d'/' -f1)))
# And the payload with the PKG_DIR prepended
PKG_PAYLOAD := $(addprefix $(PKG_DIR)/, $(PAYLOAD))

# OS detection & changes
UNAME := $(shell uname)
ifeq ($(UNAME), Linux)
  MD5BIN=md5sum
endif
ifeq ($(UNAME), Darwin)
  MD5BIN=md5
endif
ifeq ($(UNAME), FreeBSD)
  MD5BIN=md5
endif

.PHONY: all clean

all: $(PACKAGE)

$(PACKAGE): $(PAYLOAD)
	@mkdir $(PKG_DIR)
	@cp -r $(PAYLOAD_TOPS) $(PKG_DIR)
	@zip -qrX $(PACKAGE) $(PKG_PAYLOAD)
	@$(MD5BIN) $(PACKAGE)
	@rm -rf $(PKG_DIR) $(LEGAL_DOCS) $(LICENSE)

$(LICENSE): CCSI_TE_LICENSE.txt 
	@sed "s/\[SOFTWARE NAME \& VERSION\]/$(PRODUCT) v.$(VERSION)/" < CCSI_TE_LICENSE.txt > $(LICENSE)

$(LEGAL_DOCS):
	@if [ -d $(COMMON) ]; then \
	  cp $(COMMON)/$@ .; \
	else \
	  svn -q export ^/projects/common/trunk/$@; \
	fi

clean:
	@rm -rf $(PACKAGE) $(PKG_DIR) $(LEGAL_DOCS) $(LICENSE)

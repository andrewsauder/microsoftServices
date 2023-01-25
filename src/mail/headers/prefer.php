<?php

namespace andrewsauder\microsoftServices\mail\headers;

enum prefer: string {
	case HTML = 'html';
	case TEXT = 'text';
}
package com.serdarfirlayis.poimswordexample.controller;

import com.serdarfirlayis.poimswordexample.service.PoiMsWordExampleService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequiredArgsConstructor
@RequestMapping("/")
public class PoiMsWordExampleController {

    private final PoiMsWordExampleService poiMsWordExampleService;

    @GetMapping("/createWordDocument")
    public ResponseEntity createWordDocument() {
        poiMsWordExampleService.createWordDocument();
        return ResponseEntity.ok().build();
    }
}

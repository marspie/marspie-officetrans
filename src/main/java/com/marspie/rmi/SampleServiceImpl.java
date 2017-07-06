package com.marspie.rmi;

import org.springframework.stereotype.Service;

@Service("sampleServiceImpl")
public class SampleServiceImpl implements SampleService {

	@Override
	public int plus(int paramInt1, int paramInt2) {
		return paramInt1 + paramInt2;
	}

}

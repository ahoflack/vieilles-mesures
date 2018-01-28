package com.ah.cs.vieillesmesures_ear.service;

import java.util.List;
import java.util.logging.Logger;

import javax.ejb.Stateless;
import javax.enterprise.event.Event;
import javax.inject.Inject;
import javax.persistence.EntityManager;

import com.ah.cs.vieillesmesures_ear.model.MemberFile;

// The @Stateless annotation eliminates the need for manual transaction demarcation
@Stateless
public class MemberFileRegistration {

    @Inject
    private Logger log;


    @Inject
    private Event<MemberFile> memberEventSrc;
    
    @Inject
    public List<MemberFile> members;

    public void register() throws Exception {
    	
    	members.forEach(member -> memberEventSrc.fire(member) );
        log.info("Chargement terminé : " + members.size() + " fichiers");
    }
}
